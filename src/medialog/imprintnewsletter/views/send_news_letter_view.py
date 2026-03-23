# -*- coding: utf-8 -*-

## Brevo installs
# from __future__ import print_function
# import time
# import brevo_python
# from brevo_python.rest import ApiException
# from pprint import pprint

# from medialog.imprintnewsletter import _
from Acquisition import aq_inner
from datetime import date
from email import encoders
from email.message import EmailMessage
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from plone import api
from plone.registry.interfaces import IRegistry
from premailer import transform
from Products.CMFCore.utils import getToolByName
from Products.CMFPlone import PloneMessageFactory as _
from Products.CMFPlone.interfaces import IMailSchema
from Products.CMFPlone.utils import getSite
from Products.Five.browser import BrowserView
from Products.statusmessages.interfaces import IStatusMessage
from zope.annotation.interfaces import IAnnotations
from zope.component import adapter, getUtility
from zope.component import getMultiAdapter
from zope.component import getUtility
from zope.i18nmessageid import MessageFactory
from zope.interface import Interface
from zope.interface.interfaces import ComponentLookupError
import transaction
import smtplib
from medialog.imprintnewsletter import _
from medialog.imprintnewsletter.interfaces import IMedialogImprintNewsletterSettings
from medialog.imprintnewsletter.utils import get_subscriber_emails
from medialog.imprintnewsletter.views.news_letter_view import NewsLetterView
import requests



_ = MessageFactory("plone")

class ISendNewsLetterView(Interface):
    """ Marker Interface for ISendNewsLetterView"""


class SendNewsLetterView(BrowserView):
    # If you want to define a template here, please remove the template from
    # the configure.zcml registration of this view.
    # template = ViewPageTemplateFile('send_news_letter_view.pt')
    

    def __call__(self):
        request = self.request
        if 'groupmail' in request.form:
            self.send_groupmail()
        else:
            self.send_testmail() 
        # return self.index()
        # return self.request.response.redirect(self.context.absolute_url())
        
        
    
    @property
    def footer_text(self):
        return api.portal.get_registry_record('footer_text', interface=IMedialogImprintNewsletterSettings)
    
    @property
    def disclaimer_text(self):
        return api.portal.get_registry_record('disclaimer_text', interface=IMedialogImprintNewsletterSettings)
    

    def construct_message(self):
        context = self.context
        # request = self.request
        # self.email_charset = self.mail_settings.email_charset        
        title = context.Title()
        description = context.Description()
        portal_title = NewsLetterView.portal_title(self)
        navigation_root_url = NewsLetterView.navigation_root_url(self)
        img_src = NewsLetterView.get_logo(self)
        footer_text =  ""
        if self.footer_text:
            footer_text = self.footer_text.output 
        disclaimer_text = ""
        if self.disclaimer_text:
            disclaimer_text = self.disclaimer_text.output
        
        message =  u"""<html>
        <style>.text-start {text-align: left}
            .text-end {text-align: right}
            .text-center {text-align: center}
            .text-decoration-none {text-decoration: none}
            .text-decoration-underline {text-decoration: underline}
            .text-decoration-line-through {text-decoration: line-through}
            .text-lowercase {text-transform: lowercase}
            .text-uppercase {text-transform: uppercase}
            .text-capitalize {text-transform: capitalize}
            .text-wrap {white-space: normal}
            .text-nowrap {white-space: nowrap}
            .text-break {word-wrap: break-word;	word-break: break-word;}
        </style>"""
        message  += f"""<body style="margin: 0; padding: 0; font-family: Roboto, Arial, sans-serif; background-color: #f4f4f4;">
                        <div style="max-width: 600px; margin: 20px auto; 
                            background-color: #ffffff; padding: 20px; 
                            font-size: 15px; line-height: 1.6; color: #333;">                   
                            <a id="logo"
                                title="{portal_title}"
                                href="{navigation_root_url}"
                                title="{portal_title}"
                                style="max-width: 600px;"
                            >
                                <img alt="{portal_title}"
                                    title="{portal_title}"
                                    src="{img_src}"
                                    style="max-width: 600px; height: auto"
                                />
                            </a>
                            <div style="color: #555; padding: 2rem 0; margin: 2rem 0;"><hr/></div>
                            <h1 style="color: #D62265; 
                                font-weight: 400 !important;
                                font-size: 34px; margin-top: 0;">
                                {title}
                            </h1>
                            <div style="font-style: italic; color: #555; margin-bottom: 20px; font-size: 20px">
                                {description}
                            </div>
                            {context.text.output if context.text else ''}
                            <div style="color: #555; padding: 2rem 0; margin: 2rem 0;"><hr/></div>
                            
                        
                """
        message += self.more_message()
        message += footer_text 
        message +=  f"""</div>
                <div style="max-width: 600px; margin: 10px auto;">
                    {disclaimer_text} 
                </div>                
                </html>"""
        
        return transform(message)
    
    def send_groupmail(self):
        context = self.context
        request = self.request
        api_key =  api.portal.get_registry_record('api_key', interface=IMedialogImprintNewsletterSettings)
        # To do: get users from somewhere
        # Currently, it is sending to 'everybody/all_users'
        if hasattr(context, 'group'):
            group = context.group
            usergroup = api.user.get_users(groupname=group)
            
            mail_list = []
            for member in usergroup:
                group = api.group.get_groups(user=member)
                recipient = member.getProperty('email')
                # fullname = member.getProperty('fullname')
                mail_list.append(recipient) 
                # self.send_emails(context, request, recipient, fullname)
            if api_key:
                self.send_emails(context, request, mail_list, api_key)
            else: 
                self.send_emails_locally(context, request, mail_list)

        else:
            # alternatively, send to all users on site as well
            # usergroup = api.user.get_users()
            
            site = getSite()
            mail_list = get_subscriber_emails(site)

            if api_key:
                self.send_emails(context, request, mail_list, api_key)
            else: 
                self.send_emails_locally(context, request, mail_list)


        self.request.response.redirect(self.context.absolute_url())
        

    def more_message(self):
        items = NewsLetterView.get_items(self)
        if not items:
            return ''

        html_output = ''
        for obj in items:
            # obj = item.getObject()
            scales = getMultiAdapter((obj, self.request), name="images")
            thumbnail = scales.scale('image', width=600)

            image_html = ''
            if thumbnail:
                image_html = f"""
                <div style="padding: 0; margin: 0.5rem 0">
                    <figure style="padding: 0; margin:0">
                        <img style="margin: 1rem 0 0.5rem" 
                             src="{thumbnail.url}" width="{thumbnail.width}" height="{thumbnail.height}" loading="lazy"/>
                        <figcaption style="color: #777;">{obj.image_caption or ''}</figcaption>
                    </figure>
                </div>
                """

            html_output += f"""
            <article>
                {image_html}
                <a href="{obj.absolute_url()}" style="text-decoration: none">
                    <h3 style="color: #375d9a; line-height; margin-top: 0; margin-bottom: .5rem; line-height: 1.2;font-size: 30px; font-weight: 300;">{obj.Title()}</h3>
                </a>
                <p class="lead documentDescription" style="font-size: 18px; border-bottom: 1px solid #0095CA !important;
                color: #D62265 !important;
                padding-bottom: 0.5em;
                margin-bottom: 1em;
                font-weight: 200 !important;">{obj.Description()}</p>
                <div>{obj.text.output if obj.text else ''}</div>"""
                
            if obj.portal_type == 'Proloog':
                html_output += f"""
                    <p><b>Startdatum:</b> {obj.startdatum.strftime('%d-%m-%Y')}</p>"""
            
            html_output += f"""   
                <a href="{obj.absolute_url()}"
                   style="color: #fff; background-color: #D62265;  
                   border: 1px solid #D62265; padding: 0.55rem 1rem; 
                   font-size: 1.2rem; line-height: 1.75; 
                   text-decoraration: none;
                   border-radius: 0.175rem">Lees verder</a>
            </article>
            <div style="padding: 2rem 0; margin: 1rem 0;"><hr/></div>
            """
        
        return html_output


    def send_email(self, context, request, recipient, fullname):    
        registry = getUtility(IRegistry)
        self.mail_settings = registry.forInterface(IMailSchema, prefix="plone")
        #interpolator = IStringInterpolator(obj)

        mailhost = getToolByName(aq_inner(self.context), "MailHost")
        if not mailhost:
            abc = 1
            raise ComponentLookupError(
                "You must have a Mailhost utility to \
            execute this action"
            )

        # ready to create multipart mail
        try:
            self.email_charset = self.mail_settings.email_charset        
            title = context.Title()
            # description = context.Description()
            messages = IStatusMessage(self.request)
            message = self.construct_message()
            outer = MIMEMultipart('alternative')
            outer['Subject'] =  title                    
            outer['To'] = formataddr((fullname, recipient))
            newsletterfrom =  api.portal.get_registry_record('newsletter_from', interface=IMedialogImprintNewsletterSettings)
            outer['From'] =  formataddr((self.mail_settings.email_from_name, newsletterfrom))            
            
            outer.epilogue = ''

            # Attach text part
            html_part = MIMEMultipart('related')
            html_text = MIMEText(message, 'html', _charset='UTF-8')
            html_part.attach(html_text)
            outer.attach(html_part)
            # # Finally send mail.
            mailhost.send(outer.as_string())              
            
            # if we want 'not html we can use api portal
            # api.portal.send_email(
            #     recipient = formataddr((fullname, recipient)),
            #     sender = formataddr((self.mail_settings.email_from_name, self.mail_settings.email_from_address)),         
            #     subject = title, 
            #     body = message
            # )                       

            messages.add(_("sent_mail_message",  default=u"Sent to  $email",
                                                 mapping={'email': recipient },
                                                 ),
                                                 type="info")

        
        # except ConnectionRefusedError: 
        #     messages.add("Please check Email setup", type="error")        
        
        except:
            messages.add(_("cant_send_mail_message",
                                                 default=u"Could not send to $email",
                                                 mapping={'email': recipient },
                                                 ),
                                                 type="warning")
         


    # Old code  for 'own mail server'
    # Send group mail
    def send_emails_locally(self, context, request, recipients):
        registry = getUtility(IRegistry)
        self.mail_settings = registry.forInterface(IMailSchema, prefix="plone")
        #  mailhost = getToolByName(aq_inner(self.context), "MailHost")
        mailhost = getToolByName(context, "MailHost")
        messages = IStatusMessage(request)

        annotations = IAnnotations(context)
        SENT_KEY = "sent_data"

        # Load sent data
        sent_data = annotations.get(SENT_KEY, {})
        today_str = date.today().isoformat()

        # Get already sent list for today
        already_sent = sent_data.get(today_str, [])

        # Filter recipients
        recipients_to_send = [r for r in recipients if r not in already_sent]
        if not recipients_to_send:
            messages.add(_("All recipients have already received the mail today."), type="info")
            return

        title = context.Title()
        message = self.construct_message()
        smtp_host = self.mail_settings.smtp_host
        smtp_port = self.mail_settings.smtp_port
        newsletterfrom =  api.portal.get_registry_record('newsletter_from', interface=IMedialogImprintNewsletterSettings)

        try:
            
            for recipient in recipients_to_send:
                msg = EmailMessage()
                msg['Subject'] = title
                msg['From'] = formataddr((self.mail_settings.email_from_name,  newsletterfrom))
                msg['To'] = formataddr((recipient['email'], recipient['email']))
                msg.add_alternative(message, subtype='html')                
                
                if not recipient in already_sent:
                    with smtplib.SMTP(smtp_host, smtp_port) as server:
                        server.sendmail(
                            from_addr= newsletterfrom,
                            to_addrs=[recipient['email']],
                            msg=msg.as_string()
                        )

                    # Update sent record
                    already_sent.append(recipient)
                    # Save sent data back to annotations
                    sent_data[today_str] = already_sent
                    annotations[SENT_KEY] = sent_data
                    context._p_changed = True  # mark as modified for persistence
                    transaction.commit()
                
                    # # Add message to user UI
                    # messages.add(
                    #     _("sent_mail_message",
                    #     default=u"Sent to $email",
                    #     mapping={'email': recipient}),
                    #     type="info"
                    # )
            messages.add(
                        _("sent_mail_message",
                        default=u"Sent to all",
                        type="info")
                    )
                
                # sent_data[today_str] = already_sent
                # annotations[SENT_KEY] = sent_data
                # context._p_changed = True  # mark as modified for persistence
                
                #https://www.smtpeter.com


        except Exception as e:
            messages.add(
                _("cant_send_mail_message",
                default=u"Could not send to all: ${error}",
                mapping={'error': str(e)}),
                type="warning"
            )
    

    def send_emails(self, context, request, recipients, api_key):
        import pdb; pdb.set_trace()
        registry = getUtility(IRegistry)
        self.mail_settings = registry.forInterface(IMailSchema, prefix="plone")
        messages = IStatusMessage(request)
        annotations = IAnnotations(context)
        SENT_KEY = "sent_data"         
        sent_data = annotations.get(SENT_KEY, {})
        today_str = date.today().isoformat()
        already_sent = sent_data.get(today_str, [])

        recipients_to_send = [r for r in recipients if r not in already_sent]
        if not recipients_to_send:
            messages.add(_("All recipients have already received the mail today."), type="info")
            return

        title = context.Title()
        html_message = self.construct_message()

        newsletterfrom = api.portal.get_registry_record(
            'newsletter_from',
            interface=IMedialogImprintNewsletterSettings
        )

        api_url = f"https://www.smtpeter.com/v1/send?access_token={api_key}"

        try:
            payload = {
                "from": newsletterfrom,
                "to": newsletterfrom,  # visible header (safe default)
                "subject": title,
                "html": html_message,
                "text": "This email requires HTML support",
                "recipients": recipients_to_send  # 🔥 all recipients here
            }

            response = requests.post(
                api_url,
                json=payload,
                headers={"Content-Type": "application/json"}
            )

            if response.status_code not in (200, 201, 202):
                raise Exception(f"SMTPeter error: {response.text}")

            result = response.json()

            # ✅ mark all as sent
            already_sent.extend(recipients_to_send)

            sent_data[today_str] = already_sent
            annotations[SENT_KEY] = sent_data
            context._p_changed = True
            transaction.commit()

            # UI messages
            for recipient in recipients_to_send:
                messages.add(
                    _("sent_mail_message",
                    default=u"Sent to $email",
                    mapping={'email': recipient}),
                    type="info"
                )

        except Exception as e:
            messages.add(str(e), type="error")
 


    def send_testmail(self):
            context = self.context
            request = self.request
            member = api.user.get_current()
            recipient = member.getProperty('email')
            fullname = member.getProperty('fullname')
            if recipient:
                # self.send_with_brevo(context, request, recipient, fullname)
                self.send_email(context, request, recipient, fullname)
            else:
                messages = IStatusMessage(self.request)
                messages.add(_("cant_send_mail_message",
                                                    default=u"User does not have email",
                                                    mapping={'email': recipient },
                                                    ),
                                                    type="error")
                
            self.request.response.redirect(self.context.absolute_url())
            
            
            
            
    # #Use brevo api to send email
    # def send_with_brevo(self, context, request, recipient, fullname): 
    #     # Configure API key authorization: api-key
    #     configuration = brevo_python.Configuration()
    #     configuration.api_key['api-key'] = API_KEY
    #     # Uncomment below to setup prefix (e.g. Bearer) for API key, if needed
    #     # configuration.api_key_prefix['api-key'] = 'Bearer'
    #     # Configure API key authorization: partner-key
    #     # configuration = brevo_python.Configuration()
    #     # configuration.api_key['partner-key'] = API_KEY
    #     # Uncomment below to setup prefix (e.g. Bearer) for API key, if needed
    #     # configuration.api_key_prefix['partner-key'] = 'Bearer'
        
    #     ##  Message users with info
    #     messages = IStatusMessage(self.request)

    #     # create an instance of the API class
    #     api_instance = brevo_python.TransactionalEmailsApi(brevo_python.ApiClient(configuration))
    #     # subject = context.Title()
    #     subject = 'From Brevo'
    #     sender = {"name":"Brevo","email":"contact@brrevo.com"}
    #     replyTo = {"name":"Brevo","email":"contact@brevo.com"}
    #     html_content = self.construct_message()
    #     # To do, use user id and email 
    #     # outer['To'] = recipient
    #     to = [{"email":recipient,"name":fullname}]
    #     cc = [{"email":"post@medialog.no","name":"Grieg Medialog"}]
    #     bcc = [{"email":"post@medialog.no","name":"Grieg Medialog"}]
        
    #     # Not sure what this is for
    #     # params = {"parameter":"My param value","subject":"New Subject"}                    
    #     #Not sure if we need headers
    #     #headers=headers, 
        
    #     send_smtp_email = brevo_python.SendSmtpEmail(to=to, bcc=bcc, cc=cc, reply_to=replyTo,
    #                            html_content=html_content, sender=sender, subject=subject) # SendSmtpEmail | Values to send a transactional email

    #     try:
    #         # Send a transactional email
    #         api_response = api_instance.send_transac_email(send_smtp_email)
    #         # print to test if it works
    #         pprint(api_response)
    #         # give feedback to Plone
    #         # To Do
    #     except ApiException as e:
    #         print("Exception when calling TransactionalEmailsApi->send_transac_email: %s\n" % e)
            
            








###########################################

# Update contact
# configuration = sib_api_v3_sdk.Configuration()
# configuration.api_key['api-key'] = 'YOUR_API_KEY'
# api_instance = sib_api_v3_sdk.ContactsApi(sib_api_v3_sdk.ApiClient(configuration))


# create_contact = sib_api_v3_sdk.CreateContact(email="example@example.com", update_enabled = True , attributes= {'FNAME': 'JON', 'LNAME':'DOE'} ,  list_ids=[1])

# try:
#     api_response = api_instance.create_contact(create_contact)
#     print(api_response)
# except ApiException as e:
#     print("Exception when calling ContactsApi->create_contact: %s\n" % e)



# our SMTP Settings
# SMTP Server
# smtp-relay.brevo.com
# Port
# 587
# Login
# xxxxxx@smtp-brevo.com
# xxxx = fredag 4 juni skjermbilde
