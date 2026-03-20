# -*- coding: utf-8 -*-

# from medialog.imprintnewsletter import _
from Products.Five.browser import BrowserView
from zope.interface import Interface
from zope.component import getMultiAdapter
from Products.CMFPlone.utils import getSiteLogo
from medialog.imprintnewsletter.interfaces import IMedialogImprintNewsletterSettings
from plone import api
from plone.app.uuid.utils import uuidToCatalogBrain

# from Products.Five.browser.pagetemplatefile import ViewPageTemplateFile


class INewsLetterView(Interface):
    """ Marker Interface for INewsLetterView"""


class NewsLetterView(BrowserView):
    # If you want to define a template here, please remove the template from
    # the configure.zcml registration of this view.
    # template = ViewPageTemplateFile('news_letter_view.pt')

    def __call__(self):
        # Implement your own actions:
        self.items = self.get_items()
        return self.index()
    
    
    def get_logo(self):
        return getSiteLogo()
    
    @property
    def footer_text(self):
        return api.portal.get_registry_record('footer_text', interface=IMedialogImprintNewsletterSettings)
    
        
    # def get_logo_title(self):
    #     registry = getUtility(IRegistry)
    #     settings = registry.forInterface(ISiteSchema, prefix="plone", check=False)
    #     self.logo_title = settings.site_title
    #     self.img_src = getSiteLogo()
        
    # self.navigation_root_title = self.portal_state.navigation_root_title()

        
    
    def portal_url(self):
        portal_state = getMultiAdapter((self.context, self.request), name='plone_portal_state')
        return portal_state.portal_url()

    def navigation_root_url(self):
        portal_state = getMultiAdapter((self.context, self.request), name='plone_portal_state')
        return portal_state.navigation_root_url()

    def portal_title(self):
        portal_state = getMultiAdapter((self.context, self.request), name='plone_portal_state')
        return portal_state.portal_title()

    
    def get_items(self):
        context = self.context
        # TO DO, cache (?)
        # today = datetime.now()
        # could consider count from 'last month or similar'
        count = context.itemcount
        
        brains = context.portal_catalog(
            portal_type=['News Item'],
            review_state=['Published', 'published'],
            sort_on='effective',
            sort_order='descending'
        )[:count]
        
        # convert brains to objects, filter by permission
        # Might not need this, since we only search for 'Published'
        items = [
            b.getObject()
            for b in brains
            if api.user.has_permission('View', obj=b.getObject())
        ]
        
        # Resolve related items (RelationList)
        related_items = getattr(context, "related_items", None)
        if related_items:
            for rel in related_items:
                if rel.isBroken():
                    # skip broken relations
                    continue
                obj = rel.to_object
                # filter by permission
                if api.user.has_permission('View', obj=obj):
                    if obj not in items:
                        items.append(obj)        
        
        if items:
            return items
        
        # No Items found
        return None
        
