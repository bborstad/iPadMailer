#!/usr/bin/env python
# coding: utf-8

# In[8]:


#
# iPadMailer 1.2 by Carson Hunt and Benjamin Borstad for St. Jude Children's Research Hosptial - IS Logistics
# Last update: 3/19/2020
#

import win32com.client as win32
import pandas as pd
import codecs

def iPadMailer():
    def import_template():
        file = codecs.open("Template.htm", 'r')
        return file.read()
    
    def read_ss():
        df = pd.read_excel('iPadMailer.xlsx')
        return df.set_index(df.index).T.to_dict('list') #returns key: [asset, model, po, name, email]
        
    def split_name(name):
        names = name.split(' ')
        return names[0]
        
    def modify_template(template, f_name, asset, model, po, name):
        return template.replace('[f_name]', f_name).replace('[asset]', str(asset)).replace('[model]', model).replace('[po]', str(po)).replace('[name]', name)
        
    def create_email(keyed_template, asset_tag, email_to):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.SentOnBehalfOfName = 'ServiceDesk@stjude.org'
        mail.To = email_to
        mail.CC = 'CS LEAD; Enterprise Telecom Team; IS Logistics; IS EI CS HelpDesk;' #Edit this line to update CC
        mail.Subject = 'iPad Deployment ' + asset_tag
        mail.HTMLBody = keyed_template + asset_tag
        mail.Display()

    def send_emails():
        template = import_template()
        emails = read_ss()
        for key in emails:
            asset = emails[key][0]
            model = emails[key][1]
            po = emails[key][2]
            name = emails[key][3]
            f_name = split_name(name) 
            email = emails[key][4]

            #Modify template:
            keyed_template = modify_template(template, f_name, asset, model, po, name)

            #Create and send email:
            create_email(keyed_template, asset, email)
    send_emails()
iPadMailer()

