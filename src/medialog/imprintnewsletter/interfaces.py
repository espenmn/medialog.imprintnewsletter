# -*- coding: utf-8 -*-
"""Module where all interfaces, events and exceptions live."""

from medialog.controlpanel.interfaces import IMedialogControlpanelSettingsProvider
from medialog.imprintnewsletter import _
from plone.app.registry.browser.controlpanel import RegistryEditForm
from plone.app.textfield import RichText
from plone.app.textfield import RichText
from plone.app.z3cform.widgets.richtext import RichTextFieldWidget
from plone.autoform import directives
from plone.autoform.directives import widget
from plone.registry.field import PersistentField
from plone.supermodel import model 
from z3c.form import interfaces
from zope import schema
from zope.interface import alsoProvides
from zope.interface import Interface
from zope.publisher.interfaces.browser import IDefaultBrowserLayer




class IMedialogImprintNewsletterLayer(IDefaultBrowserLayer):
    """Marker interface that defines a browser layer."""


class RichTextFieldRegistry(PersistentField, RichText):
    """ persistent registry textfield """


def richtextConstraint(value):
    """ Workaround for bug 
    """
    value = value.output
    return True 

def richtextget(value):
    """ Workaround for bug 
    """
    value = value.output
    return value 


class IMedialogImprintNewsletterSettings(model.Schema):
    """Adds settings to medialog.controlpanel
    """

    model.fieldset(
        'newsletter',
        label=_(u'Newsletter'),
        fields=[
            'footer_text',
            'disclaimer_text',
            'newsletter_from',
            'api_key'
            ],
        )
    
    widget(footer_text=RichTextFieldWidget)
    footer_text = RichTextFieldRegistry(
        title="Footer text",
        required=False,
    )
    
    widget(disclaimer_text=RichTextFieldWidget)
    disclaimer_text = RichTextFieldRegistry(
        title="Disclaimer text",
        required=False,
    )
    
    newsletter_from = schema.TextLine(
        title="Newsletter email from address",
        required=True,
    )
    
    api_key = schema.TextLine(
        title="SMPETER API KEY",
        required=True,
    )

alsoProvides(IMedialogImprintNewsletterSettings, IMedialogControlpanelSettingsProvider)