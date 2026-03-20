# -*- coding: utf-8 -*-
# import csv
import io
import requests
# from plone import api
from Products.CMFPlone.utils import getSite
from Products.Five import BrowserView
from zope.interface import Interface
#from zope.schema import Bytes
from z3c.form import form, field, button
from plone.namedfile.field import NamedBlobFile 
from zope.annotation.interfaces import IAnnotations
from persistent.list import PersistentList

from Products.CMFPlone.utils import getSite
from io import BytesIO
import pandas as pd

SUBSCRIBERS_KEY = 'medialog.imprintnewsletter.subscribers'


# NOTE this code does not work for Form (only BrowserView)

class ExportNewsletterEmails(BrowserView):

    def __call__(self):
        site = getSite()
        annotations = IAnnotations(site)
        storage = annotations.get(SUBSCRIBERS_KEY, [])
        df = pd.DataFrame({"email": list(storage)})
        output = BytesIO()
        df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)

        response = self.request.response
        response.setHeader(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        response.setHeader(
            "Content-Disposition",
            'attachment; filename="newsletter_emails.xlsx"'
        )

        return output.getvalue()

 

class IEmailListImport(Interface):
    """ Marker Interface for IMeetingImport"""

class IExcelImportFormSchema(Interface):
    
    excel_file = NamedBlobFile(
        title=u"Excel File",
        description=u"Upload a Excel file with email list for newsletter",
        required=False
    )
 

class EmailList(form.Form):
    fields = field.Fields(IExcelImportFormSchema)
    ignoreContext = True
    label = u"Import Newsletter email list from Excel to replace existing emails"

    @button.buttonAndHandler(u"Import")
    def handleImport(self, action):
        data, errors = self.extractData()
        if errors:
            self.status = self.formErrorsMessage
            return

        file_data = data['excel_file']

        # Get raw bytes from the upload
        if hasattr(file_data, 'data'):                 # plone.namedfile NamedBlobFile
            raw = file_data.data
        elif hasattr(file_data, 'read'):               # ZPublisher FileUpload
            raw = file_data.read()
        elif isinstance(file_data, (bytes, bytearray)):
            raw = bytes(file_data)
        else:
            self.status += f"Wrong or missing Excel file"
            return None
            # raise ValueError("Unsupported upload type for Excel file")
        

        # Read the sheet as strings for the rest of the data
        df = pd.read_excel(BytesIO(raw), sheet_name=0, dtype=str)
        df = df.fillna("")
        df.columns = [str(c).strip().lower() for c in df.columns]
        rows = df.to_dict(orient="records")

        storage = self._get_storage()

        # Remove duplicates
        userlist = {row.get("email") for row in rows if row.get("email")}

        # Modify persistent list in-place
        storage.clear()
        storage.extend(userlist)
        self.status = f"Imported {len(userlist)} Email addresses"
         
        

    def _get_storage(self):    
        site = getSite()
        annotations = IAnnotations(site)
        if SUBSCRIBERS_KEY not in annotations:
            annotations[SUBSCRIBERS_KEY] = PersistentList()
        return annotations[SUBSCRIBERS_KEY]

 