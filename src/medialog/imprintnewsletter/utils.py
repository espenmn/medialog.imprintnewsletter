from zope.annotation.interfaces import IAnnotations
from persistent.list import PersistentList


SUBSCRIBERS_KEY = 'medialog.imprintnewsletter.subscribers'

def get_subscriber_emails(context):
    annotations = IAnnotations(context)
    return annotations.get(SUBSCRIBERS_KEY, PersistentList())
