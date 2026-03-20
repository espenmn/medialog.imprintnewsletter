# ============================================================================
# DEXTERITY ROBOT TESTS
# ============================================================================
#
# Run this robot test stand-alone:
#
#  $ bin/test -s medialog.imprintnewsletter -t test_news_letter.robot --all
#
# Run this robot test with robot server (which is faster):
#
# 1) Start robot server:
#
# $ bin/robot-server --reload-path src medialog.imprintnewsletter.testing.MEDIALOG_NEWSLETTER_ACCEPTANCE_TESTING
#
# 2) Run robot tests:
#
# $ bin/robot /src/medialog/newsletter/tests/robot/test_news_letter.robot
#
# See the http://docs.plone.org for further details (search for robot
# framework).
#
# ============================================================================

*** Settings *****************************************************************

Resource  plone/app/robotframework/selenium.robot
Resource  plone/app/robotframework/keywords.robot

Library  Remote  ${PLONE_URL}/RobotRemote

Test Setup  Open test browser
Test Teardown  Close all browsers


*** Test Cases ***************************************************************

Scenario: As a site administrator I can add a NewsLetter
  Given a logged-in site administrator
    and an add NewsLetter form
   When I type 'My NewsLetter' into the title field
    and I submit the form
   Then a NewsLetter with the title 'My NewsLetter' has been created

Scenario: As a site administrator I can view a NewsLetter
  Given a logged-in site administrator
    and a NewsLetter 'My NewsLetter'
   When I go to the NewsLetter view
   Then I can see the NewsLetter title 'My NewsLetter'


*** Keywords *****************************************************************

# --- Given ------------------------------------------------------------------

a logged-in site administrator
  Enable autologin as  Site Administrator

an add NewsLetter form
  Go To  ${PLONE_URL}/++add++NewsLetter

a NewsLetter 'My NewsLetter'
  Create content  type=NewsLetter  id=my-news_letter  title=My NewsLetter

# --- WHEN -------------------------------------------------------------------

I type '${title}' into the title field
  Input Text  name=form.widgets.IBasic.title  ${title}

I submit the form
  Click Button  Save

I go to the NewsLetter view
  Go To  ${PLONE_URL}/my-news_letter
  Wait until page contains  Site Map


# --- THEN -------------------------------------------------------------------

a NewsLetter with the title '${title}' has been created
  Wait until page contains  Site Map
  Page should contain  ${title}
  Page should contain  Item created

I can see the NewsLetter title '${title}'
  Wait until page contains  Site Map
  Page should contain  ${title}
