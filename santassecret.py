import smtplib  
import sys
import random
from random import seed
import pandas as pd

def setInitialCondition(mate,supr):
    pos=random.randint(0,supr-2)
    mate[pos]=(supr-1)
    return pos      

def findpair(supr):
    i=0
    mate=[supr+1 for dummy in range(supr)]#
    pos=setInitialCondition(mate,supr)
    while(i<supr):
        seed()
        partner=random.randint(0,supr-1)
        if i==pos:# skip this cause already initialised
            i+=1
        if ((partner!=i) and not(partner in mate)):
            mate[i]=partner
            i+=1
    return mate
  
def sendMails(adr,namen,partner,user):
    subject = 'Wichtelpartner'
    MAIL_FROM = user 
    for i in range(0,len(namen)):
        mail_text = 'Hallo {0},\nDu hast {1} als Wichtelpartner. :)\n\n'.format(namen[i],namen[partner[i]])+'Liebe Gruesse von der Wichtelfee' # Person i gifts person j 
        RCPT_TO  = adr[i]
        DATA = 'From:%s\nTo:%s\nSubject:%s\n\n%s' % (MAIL_FROM, RCPT_TO, subject, mail_text)
        print(mail_text)
        print()
        server.sendmail(MAIL_FROM, RCPT_TO, DATA)

        
    # ################ Main
if __name__ == "__main__":
    # get userdata
    user = sys.argv[1]
    pwd = sys.argv[2]
    
    df = pd.read_excel(r'partner_data.xlsx')
    df=df.loc[:, ~df.columns.str.contains('^Unnamed')]
    df.columns = map(str.lower, df.columns) # dont mind upper and lowercase letters
    emailDict={}


    try:
        names=df["names"]
        # print(names)
    except:
        print("No column with name 'names' found")

    try:
        mail=df["emails"]
    except:
        print("No column with name 'emails' found")

    for count,mail in enumerate(df["emails"]):
        emailDict[count]=mail

    # print(emailDict)
    # create dict for list of participants
    # print(mail)
    partner=findpair(len(emailDict))
    print(partner)
    # # serverlogin
    server = smtplib.SMTP('smtp.web.de',587)
    server.starttls()
    server.login(user, pwd)
    sendMails(emailDict,names,partner,user)
    server.quit() 

