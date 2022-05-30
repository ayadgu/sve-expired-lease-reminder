import re
import glob
import pandas as pd
from datetime import datetime, timedelta, date
import smtplib, ssl
from email.message import EmailMessage
from prettytable import PrettyTable
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import sys
import os

class ReminderBot():
    def __init__(self):
        self.initVariable()
        self.initDataFrame()

    def initVariable(self):
        os.path.expanduser('~')
        print(os.environ)
        self.six_months = timedelta(6*365/12)
        self.height_months = timedelta(8*365/12)
        self.three_years = timedelta(3*365)
        self.today_date = datetime.now()

        self.PASSWORD_OUTLOOK=os.environ['PASSWORD_OUTLOOK']
        self.EMAIL_OUTLOOK=os.environ['EMAIL_OUTLOOK']
        self.EMAIL_DEST=os.environ['EMAIL_DEST']

    def initDataFrame(self):
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)
        #--------------------------
        # DF Fichier Code
        filename=glob.glob("*baux*.xlsx")[0]
        df_fichier_baux = pd.read_excel(filename)

        col_fin_bail =  'Fin Bail'
        col_type_bail =  'Type Bail'
        col_numero_bail =  'Numero Mandat'

        df_fichier_baux=df_fichier_baux.rename(columns={'20/05/2022.1': col_fin_bail})
        df_fichier_baux=df_fichier_baux.rename(columns={'Unnamed: 17': col_type_bail})
        df_fichier_baux=df_fichier_baux.rename(columns={'Unnamed: 0': col_numero_bail})

        df_fichier_baux = df_fichier_baux[df_fichier_baux[col_fin_bail].notna()]

        df_new_df=pd.concat([df_fichier_baux[col_numero_bail],df_fichier_baux[col_type_bail],df_fichier_baux[col_fin_bail]], axis=1)

        self.df_new_df = df_new_df


    def send_mail(self):
        
        text = f"""
        Ceci est un message automatique.

        Baux arrivés à expiration:

        {self.expired_bails_monthly.get_html_string()}
        {self.expired_bails_daily.get_html_string()}

        Cordialement
        """

        html = f"""
        <html><body><p>Ceci est un message automatique.</p>
        <p>Baux arrivés à expiration:</p>
        {self.expired_bails_monthly.get_html_string()}
        {self.expired_bails_daily.get_html_string()}
        <p>Cordialement</p>
        </body></html>
        """
        msg = MIMEMultipart('alternative')
        # msg = EmailMessage()
        # msg.set_content((self.expired_bails.get_html_string()))
        msg["Subject"] = "Rappel automatique - Baux bientôt arrivés à expiration"
        msg["From"] = self.EMAIL_OUTLOOK
        msg["To"] = self.EMAIL_DEST


        SMTP = "smtp-mail.outlook.com"
        context=ssl.create_default_context()

        # Record the MIME types of both parts - text/plain and text/html.
        part1 = MIMEText(text, 'plain')
        part2 = MIMEText(html, 'html')

        msg.attach(part1)
        msg.attach(part2)

        with smtplib.SMTP(SMTP, port=587) as smtp:
            smtp.starttls(context=context)
            smtp.login(self.EMAIL_OUTLOOK, self.PASSWORD_OUTLOOK)
            smtp.ehlo()

            smtp.send_message(msg)
            smtp.quit()

    def apply(self):
        
        expired_bails_monthly = PrettyTable()
        expired_bails_daily = PrettyTable()

        self.has_expired_bail_monthly=False
        self.has_expired_bail_daily=False

        expired_bails_monthly.field_names  = ["Numéro Mandat","Type Bail","Date Fin"]
        expired_bails_daily.field_names  = ["Numéro Mandat","Type Bail","Date Fin"]

        for indices, row in self.df_new_df.iterrows():
            date = self.df_new_df.at[indices,"Fin Bail"]
            numero_bail = self.df_new_df.at[indices,"Numero Mandat"]
            date_to_time_obj = datetime.strptime(date, '%d/%m/%Y')
            type_bail = str(self.df_new_df.at[indices,"Type Bail"])
            match type_bail :
                case '4':
                    if date_to_time_obj or self.today_date > date_to_time_obj + self.three_years-self.six_months :
                        self.has_expired_bail_monthly=True
                        expired_bails_monthly.add_row([numero_bail,type_bail,date])
                    
                    if self.today_date == date_to_time_obj or self.today_date == date_to_time_obj + self.three_years-self.six_months :
                        self.has_expired_bail_daily=True
                        expired_bails_daily.add_row([numero_bail,type_bail,date])
                
                case _:
                    if self.today_date > date_to_time_obj-self.height_months:
                        self.has_expired_bail_monthly=True
                        expired_bails_monthly.add_row([numero_bail,type_bail,date])
                        
                    if self.today_date == date_to_time_obj-self.height_months:
                        self.has_expired_bail_daily=True
                        expired_bails_daily.add_row([numero_bail,type_bail,date])

        self.expired_bails_monthly=expired_bails_monthly
        self.expired_bails_daily=expired_bails_daily

        print("expired_bails_daily : \n",self.expired_bails_daily)
        print("expired_bails_monthly : \n",self.expired_bails_monthly)
        print(self.today_date.day == 1, "first day")

        if (self.has_expired_bail_monthly and self.today_date.day == 1 ) or self.has_expired_bail_daily:
            print("send mail")
            self.send_mail()
        return

def main():
    print("Lancement de l'application, veuillez patienter...")
    reminder = ReminderBot()
    reminder.apply()
       
if __name__ == "__main__":
    main()
    
