# -*- coding: utf-8 -*-
from plone.app.robotframework.testing import REMOTE_LIBRARY_BUNDLE_FIXTURE
from plone.app.testing import (
    applyProfile,
    FunctionalTesting,
    IntegrationTesting,
    PLONE_FIXTURE
    PloneSandboxLayer,
)
from plone.testing import z2

import medialog.imprintnewsletter


class MedialogImprintNewsletterLayer(PloneSandboxLayer):

    defaultBases = (PLONE_FIXTURE,)

    def setUpZope(self, app, configurationContext):
        # Load any other ZCML that is required for your tests.
        # The z3c.autoinclude feature is disabled in the Plone fixture base
        # layer.
        import plone.app.dexterity
        self.loadZCML(package=plone.app.dexterity)
        import plone.restapi
        self.loadZCML(package=plone.restapi)
        self.loadZCML(package=medialog.imprintnewsletter)

    def setUpPloneSite(self, portal):
        applyProfile(portal, 'medialog.imprintnewsletter:default')


MEDIALOG_NEWSLETTER_FIXTURE = MedialogImprintNewsletterLayer()


MEDIALOG_NEWSLETTER_INTEGRATION_TESTING = IntegrationTesting(
    bases=(MEDIALOG_NEWSLETTER_FIXTURE,),
    name='MedialogImprintNewsletterLayer:IntegrationTesting',
)


MEDIALOG_NEWSLETTER_FUNCTIONAL_TESTING = FunctionalTesting(
    bases=(MEDIALOG_NEWSLETTER_FIXTURE,),
    name='MedialogImprintNewsletterLayer:FunctionalTesting',
)


MEDIALOG_NEWSLETTER_ACCEPTANCE_TESTING = FunctionalTesting(
    bases=(
        MEDIALOG_NEWSLETTER_FIXTURE,
        REMOTE_LIBRARY_BUNDLE_FIXTURE,
        z2.ZSERVER_FIXTURE,
    ),
    name='MedialogImprintNewsletterLayer:AcceptanceTesting',
)
