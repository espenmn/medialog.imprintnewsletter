# -*- coding: utf-8 -*-
from plone.app.textfield import RichText
from plone.autoform import directives
from plone.dexterity.content import Item
from plone.supermodel import model
from zope import schema
from zope.interface import implementer
from plone.app.z3cform.widget import RelatedItemsFieldWidget
from plone.app.vocabularies.catalog import CatalogSource
from z3c.relationfield.schema import RelationList, RelationChoice
# from plone.namedfile import field as namedfile
# from plone.supermodel.directives import fieldset
# from z3c.form.browser.radio import RadioFieldWidget

from medialog.imprintnewsletter import _


class INewsLetter(model.Schema):
    """ Marker interface and Dexterity Python Schema for NewsLetter
    """
   
    # model.load('news_letter.xml')
    newsletter_language = schema.Choice(
        title="Newsletter language",
        values=("all", "nl", "en", "fr", "de"),
        required=True,
    )
    
    text = RichText(
        title=_(u'Text'),
        required=False
    )
    
    itemcount = schema.Int(
        title=_(u'Number of News Items to include'),
        required=False
    )
    
    directives.widget(related_items=RelatedItemsFieldWidget)
    related_items = RelationList(
        title=_(u"Items to include"),
        default=[],
        required=False,
        value_type=RelationChoice(
            title=_(u"Content to include"),
            # source=CatalogSource(),  # optional filter here
            source=CatalogSource(portal_type=("News Item", "Document", "Proloog"))
        )
    )
    
    

    # url = schema.URI(
    #     title=_(u'Link'),
    #     required=False
    # )

    # fieldset('Images', fields=['logo', 'advertisement'])
    # logo = namedfile.NamedBlobImage(
    #     title=_(u'Logo'),
    #     required=False,
    # )

    # advertisement = namedfile.NamedBlobImage(
    #     title=_(u'Advertisement (Gold-sponsors and above)'),
    #     required=False,
    # )

    # directives.read_permission(notes='cmf.ManagePortal')
    # directives.write_permission(notes='cmf.ManagePortal')
    # notes = RichText(
    #     title=_(u'Secret Notes (only for site-admins)'),
    #     required=False
    # )


@implementer(INewsLetter)
class NewsLetter(Item):
    """ Content-type class for INewsLetter
    """
    
    def something(self):
        return 'something'
