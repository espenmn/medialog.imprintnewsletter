# -*- coding: utf-8 -*-
from medialog.imprintnewsletter.content.news_letter import INewsLetter  # NOQA E501
from medialog.imprintnewsletter.testing import MEDIALOG_NEWSLETTER_INTEGRATION_TESTING  # noqa
from plone import api
from plone.app.testing import setRoles
from plone.app.testing import TEST_USER_ID
from plone.dexterity.interfaces import IDexterityFTI
from zope.component import createObject
from zope.component import queryUtility

import unittest




class NewsLetterIntegrationTest(unittest.TestCase):

    layer = MEDIALOG_NEWSLETTER_INTEGRATION_TESTING

    def setUp(self):
        """Custom shared utility setup for tests."""
        self.portal = self.layer['portal']
        setRoles(self.portal, TEST_USER_ID, ['Manager'])
        self.parent = self.portal

    def test_ct_news_letter_schema(self):
        fti = queryUtility(IDexterityFTI, name='NewsLetter')
        schema = fti.lookupSchema()
        self.assertEqual(INewsLetter, schema)

    def test_ct_news_letter_fti(self):
        fti = queryUtility(IDexterityFTI, name='NewsLetter')
        self.assertTrue(fti)

    def test_ct_news_letter_factory(self):
        fti = queryUtility(IDexterityFTI, name='NewsLetter')
        factory = fti.factory
        obj = createObject(factory)

        self.assertTrue(
            INewsLetter.providedBy(obj),
            u'INewsLetter not provided by {0}!'.format(
                obj,
            ),
        )

    def test_ct_news_letter_adding(self):
        setRoles(self.portal, TEST_USER_ID, ['Contributor'])
        obj = api.content.create(
            container=self.portal,
            type='NewsLetter',
            id='news_letter',
        )

        self.assertTrue(
            INewsLetter.providedBy(obj),
            u'INewsLetter not provided by {0}!'.format(
                obj.id,
            ),
        )

        parent = obj.__parent__
        self.assertIn('news_letter', parent.objectIds())

        # check that deleting the object works too
        api.content.delete(obj=obj)
        self.assertNotIn('news_letter', parent.objectIds())

    def test_ct_news_letter_globally_addable(self):
        setRoles(self.portal, TEST_USER_ID, ['Contributor'])
        fti = queryUtility(IDexterityFTI, name='NewsLetter')
        self.assertTrue(
            fti.global_allow,
            u'{0} is not globally addable!'.format(fti.id)
        )
