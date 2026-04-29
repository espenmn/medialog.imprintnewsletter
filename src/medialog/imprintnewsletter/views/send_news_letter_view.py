# -*- coding: utf-8 -*-


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
                            <h1 style="color: #c92a37; 
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
                mail_list.append({
                    "email": member.getProperty('email'),
                    "language": getattr(member, "language", "en")  # fallback
                }) 
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
        
        image_width  =  api.portal.get_registry_record('image_width', interface=IMedialogImprintNewsletterSettings)
        image_height =  api.portal.get_registry_record('image_height', interface=IMedialogImprintNewsletterSettings)
        
        html_output = ''
        for obj in items:
            # obj = item.getObject()
            scales = getMultiAdapter((obj, self.request), name="images")
            thumbnail = scales.scale('image', width=image_width, height=image_height)

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
                    <h3 style="color: #21579f; line-height; margin-top: 0; margin-bottom: .5rem; line-height: 1.2;font-size: 30px; font-weight: 300;">{obj.Title()}</h3>
                </a>
                <p class="lead documentDescription" style="font-size: 18px; border-bottom: 1px solid #0095CA !important;
                color: #c92a37 !important;
                padding-bottom: 0.5em;
                margin-bottom: 1em;
                font-weight: 200 !important;">{obj.Description()}</p>"""
                # <div>{obj.text.output if obj.text else ''}</div>"""
                
            if obj.portal_type == 'Proloog':
                html_output += f"""
                    <p><b>Startdatum:</b> {obj.startdatum.strftime('%d-%m-%Y')}</p>"""
            
            html_output += f"""   
                <a href="{obj.absolute_url()}"
                   style="color: #fff; background-color: #c92a37;  
                   border: 1px solid #c92a37; padding: 0.55rem 1rem; 
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
        mailhost = getToolByName(context, "MailHost")
        messages = IStatusMessage(request)

        annotations, sent_data, today_str, already_sent = self._get_sent_state(context)

        recipients_to_send = self._get_recipients_to_send(context, recipients, already_sent)

        if not recipients_to_send:
            messages.add(_("All recipients have already received the mail today."), type="info")
            return

        title = context.Title()
        message = self.construct_message()
        smtp_host = self.mail_settings.smtp_host
        smtp_port = self.mail_settings.smtp_port

        newsletterfrom = api.portal.get_registry_record(
            'newsletter_from',
            interface=IMedialogImprintNewsletterSettings
        )

        sent_emails = []

        try:
            for r in recipients_to_send:
                msg = EmailMessage()
                msg['Subject'] = title
                msg['From'] = formataddr((self.mail_settings.email_from_name, newsletterfrom))
                msg['To'] = r["email"]
                msg.add_alternative(message, subtype='html')

                with smtplib.SMTP(smtp_host, smtp_port) as server:
                    server.sendmail(
                        from_addr=newsletterfrom,
                        to_addrs=[r["email"]],
                        msg=msg.as_string()
                    )

                sent_emails.append(r["email"])

            self._mark_as_sent(context, annotations, sent_data, today_str, already_sent, sent_emails)

            messages.add(
                _("sent_to_message", default="Sent to ${count}.", mapping={"count": len(sent_emails)}),
                type="info"
            )

        except Exception as e:
            messages.add(
                _("cant_send_mail_message",
                default=u"Could not send: ${error}",
                mapping={'error': str(e)}),
                type="warning"
            )
    

    def send_emails(self, context, request, recipients, api_key):
        messages = IStatusMessage(request)

        annotations, sent_data, today_str, already_sent = self._get_sent_state(context)

        recipients_to_send = self._get_recipients_to_send(context, recipients, already_sent)

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
                "to": newsletterfrom,
                "subject": title,
                "html": html_message,
                "text": "This email requires HTML support",
                "recipients": [r["email"] for r in recipients_to_send]
            }

            response = requests.post(api_url, json=payload)

            if response.status_code not in (200, 201, 202):
                raise Exception(response.text)

            sent_emails = [r["email"] for r in recipients_to_send]

            self._mark_as_sent(context, annotations, sent_data, today_str, already_sent, sent_emails)

            messages.add(
                _("sent_mail_message",
                default=u"Sent to: $emails",
                mapping={'emails': ", ".join(sent_emails)}),
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
            
            
             
             
    def _get_sent_state(self, context):
        annotations = IAnnotations(context)
        SENT_KEY = "sent_data"

        sent_data = annotations.get(SENT_KEY, {})
        today_str = date.today().isoformat()
        already_sent = sent_data.get(today_str, [])

        return annotations, sent_data, today_str, already_sent


    def _mark_as_sent(self, context, annotations, sent_data, today_str, already_sent, new_emails):
        already_sent.extend(new_emails)
        sent_data[today_str] = already_sent
        annotations["sent_data"] = sent_data
        context._p_changed = True
        transaction.commit()


    def _get_recipients_to_send(self, context, recipients, already_sent):
        newsletter_language = getattr(context, "newsletter_language", "all")

        return [
            r for r in recipients
            if (
                r.get("email")
                and r["email"] not in already_sent
                and (
                    newsletter_language == "all"
                    or r.get("language") == newsletter_language
                )
            )
        ]