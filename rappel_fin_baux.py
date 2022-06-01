import re
import glob
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, date
import smtplib, ssl
from email.message import EmailMessage
from prettytable import PrettyTable
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import sys
import os
import time
from dotenv import load_dotenv


class ReminderBot():
    def __init__(self):
        self.initVariable()
        self.initDataFrame()

    def initVariable(self):
        load_dotenv()
        os.path.expanduser('~')

        self.six_months = timedelta(6*365/12)
        self.height_months = timedelta(8*365/12)
        self.three_years = timedelta(3*365)
        self.today_date = datetime.today()

        self.PASSWORD_OUTLOOK=os.environ.get("PASSWORD_OUTLOOK")
        self.EMAIL_OUTLOOK=os.environ.get("EMAIL_OUTLOOK")
        self.EMAIL_DEST=os.environ.get("EMAIL_DEST")

        

    def initDataFrame(self):
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)

        #--------------------------
        # DF Fichier Adresse des baux
        filename=glob.glob("*adresse*.xlsx")[0]
        df_adresse_baux = pd.read_excel(filename, header= None)

        col_adresse_bail =  'Adresse Bail'
        col_numero_bail =  'Numero Mandat'

        df_adresse_baux=df_adresse_baux.rename(columns={1: col_numero_bail})
        df_adresse_baux=df_adresse_baux.rename(columns={6: col_adresse_bail})
        df_adresse_baux[col_adresse_bail] = df_adresse_baux[col_adresse_bail].astype(str)

        df_adresse_baux.replace("Indicatif",np.nan,regex=True,inplace=True)
        df_adresse_baux.replace("Entré le",np.nan,regex=True,inplace=True)
        df_adresse_baux.replace("nan",np.nan,regex=True,inplace=True)
        df_adresse_baux = df_adresse_baux[df_adresse_baux[col_adresse_bail].notna()]

        df_adresse_baux_min=pd.concat([df_adresse_baux[col_numero_bail],df_adresse_baux[col_adresse_bail]], axis=1)
        df_adresse_baux_min[col_numero_bail].ffill(inplace=True)
        df_adresse_baux_min.dropna(inplace=True)
        df_adresse_baux_min=df_adresse_baux_min.groupby([col_numero_bail], as_index=False).agg({col_adresse_bail: ' '.join})
        df_adresse_baux_min[col_adresse_bail]=df_adresse_baux_min[[col_adresse_bail,col_numero_bail]].groupby([col_numero_bail], as_index=False)[col_adresse_bail].transform(lambda x: ','.join(x))
        
        self.df_adresse_baux = df_adresse_baux_min

        #--------------------------
        # DF Fichier Situation des baux
        filename=glob.glob("*tuation*.xlsx")[0]
        df_fichier_baux = pd.read_excel(filename, header= None)

        c_time = os.path.getctime(filename)


        self.date_fichier_excel=datetime.fromtimestamp(c_time).strftime('%d/%m/%Y')
        self.path_fichier_excel=os.path.abspath(filename)
        
        col_fin_bail =  'Fin Bail'
        col_type_bail =  'Type Bail'
        col_numero_bail =  'Numero Mandat'
        col_nom_locataire =  'Nom Locataire'

        df_fichier_baux=df_fichier_baux.rename(columns={19: col_fin_bail})
        df_fichier_baux=df_fichier_baux.rename(columns={17: col_type_bail})
        df_fichier_baux=df_fichier_baux.rename(columns={0: col_numero_bail})
        df_fichier_baux=df_fichier_baux.rename(columns={3: col_nom_locataire})

        df_fichier_baux[col_type_bail] = df_fichier_baux[col_type_bail].astype(str)
        df_fichier_baux[col_type_bail]=df_fichier_baux[col_type_bail].replace(["0","1","2","3","4","5","6","7","8","9","A","B","C","D","E","F","G","H"],["0 Code Civil ICC","1 Loi Quillot 3 ans","2 Loi Quillot 6 ans","3 Meublé","4 Commercial","5 Professionnel","6 Loi de 48","7 Bail de 6 ans","8 Bail Mehaignerie 3 ans","9 Bail Mehaignerie 8 ans","A Bail Mehaignerie 6 ans","B Bail Loi 06.07.89 3 ans","C Bail Loi 06.07.89 6 ans","D Bail Loi 07.89 3 ans 1/6","E Bail comm. Dg Non Revisable","F Bail SRU","G Bail Derog. (art L.145-5 cc)","H Code Civil IRL"])


        df_fichier_baux[col_nom_locataire].bfill(inplace=True)

        df_fichier_baux = df_fichier_baux[df_fichier_baux[col_fin_bail].notna()]
        df_fichier_baux = df_fichier_baux[df_fichier_baux[col_numero_bail].notna()]

        df_situation_des_baux=pd.concat([df_fichier_baux[col_numero_bail],df_fichier_baux[col_nom_locataire],df_fichier_baux[col_type_bail],df_fichier_baux[col_fin_bail]], axis=1)
        print(df_situation_des_baux.head(5))
        
        self.df_situation_des_baux = df_situation_des_baux

        inner_join = pd.merge(df_situation_des_baux, 
                            df_adresse_baux_min, 
                            on =col_numero_bail, 
                            how ='inner')

        self.inner_join=inner_join

        

    def send_mail(self):
        text = f"""
        <h6>Ceci est un message automatique.</h6>

        <h2>Baux arrivés à expiration:</h2>

        {self.expired_bails_monthly.get_html_string()}
        {self.expired_bails_daily.get_html_string()}

        Cordialement
        """

        html = f"""
        <html>
            <body>
                <p>Ceci est un message automatique.</p>
                <h3>Baux bientôt expirés depuis aujourd'hui exactement :</h3>
                {self.expired_bails_daily.get_html_string()}
                <br/>
                <h3>Baux bientôt ou déjà expirés à ce jour :</h3>
                {self.expired_bails_monthly.get_html_string()}
                <br/>
                <p>Le fichier Excel <b>{self.path_fichier_excel}</b> utilisé pour ce rappel date du <b>{self.date_fichier_excel}</b>.</p>
                <p>
                    Pensez à le mettre à jour de temps en temps pour continuer à recevoir des rappels pertinents.
                    </br>Pour cela, passez par le menu <b>Quittancement</b> > <b>Révision des baux</b> depuis Gercop. 
                    </br>Le fichier extrait sera de type .xls, ouvrez-le puis enregistrez-le sous format .xlsx et écrasez le fichier excel actuellement utilisé.
                </p>
                <p>Bonne journée !</p>
            </body>
            </html>
        """
        text=html
        msg = MIMEMultipart('alternative')
        # msg = EmailMessage()
        # msg.set_content((self.expired_bails_daily.get_html_string()))
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
        # Init PrettyTable
        expired_bails_monthly = PrettyTable()
        expired_bails_daily = PrettyTable()

        # Naming Head
        expired_bails_monthly.field_names  = ["Numéro Mandat","Nom Locataire","Adresse Bail","Type Bail","Date Fin"]
        expired_bails_daily.field_names  = ["Numéro Mandat","Nom Locataire","Adresse Bail","Type Bail","Date Fin"]
        
        # Sorting 
        expired_bails_monthly.sortby = 'Date Fin'
        expired_bails_daily.sortby = 'Date Fin'
        # Sorting 
        expired_bails_monthly.align = 'r'
        expired_bails_daily.align = 'r'

        found_expired_lease_monthly=False
        found_expired_lease_daily=False

        for indices, row in self.inner_join.iterrows():
            date = self.inner_join.at[indices,"Fin Bail"]
            date = datetime.strptime(date, '%d/%m/%Y')
            print("")
            numero_bail = self.inner_join.at[indices,"Numero Mandat"]
            nom_locataire = self.inner_join.at[indices,"Nom Locataire"]
            adresse_locataire = self.inner_join.at[indices,"Adresse Bail"]
            type_bail = str(self.inner_join.at[indices,"Type Bail"])
            empty_row=["Aucun"]*len(self.inner_join.columns)

            match type_bail :
                # #########
                # Code bails
                # # 4 : Prévenir à la date de fin et (date de fin + 3 ans - 6 mois)
                # 0, 3, B, C, G : Prévenir 8 mois avant date de fin
                
                case '4':
                    if self.today_date == date or self.today_date > date + self.three_years-self.six_months:
                        found_expired_lease_monthly=True
                        expired_bails_monthly.add_row([numero_bail,nom_locataire,adresse_locataire,type_bail,date.strftime('%d/%m/%Y')])
                    
                    if self.today_date == date or self.today_date == date + self.three_years-self.six_months:
                        found_expired_lease_daily=True
                        expired_bails_daily.add_row([numero_bail,nom_locataire,adresse_locataire,type_bail,date.strftime('%d/%m/%Y')])
                
                case _:
                    if self.today_date > date-self.height_months:
                        found_expired_lease_monthly=True
                        expired_bails_monthly.add_row([numero_bail,nom_locataire,adresse_locataire,type_bail,date.strftime('%d/%m/%Y')])
                        
                    if self.today_date == date-self.height_months:
                        found_expired_lease_daily=True
                        expired_bails_daily.add_row([numero_bail,nom_locataire,adresse_locataire,type_bail,date.strftime('%d/%m/%Y')])

        self.expired_bails_monthly=expired_bails_monthly
        self.expired_bails_daily=expired_bails_daily

        print("expired_bails_daily : \n",self.expired_bails_daily)
        print("expired_bails_monthly : \n",self.expired_bails_monthly)
        print(self.today_date.day == 1, "first day")

        print("found_expired_lease_monthly",found_expired_lease_monthly)
        print("found_expired_lease_daily",found_expired_lease_daily)

        if not found_expired_lease_daily :
            self.expired_bails_daily.add_row(empty_row)

        if not found_expired_lease_monthly :
            self.expired_bails_monthly.add_row(empty_row)

        print(self.date_fichier_excel)

        if (found_expired_lease_monthly and self.today_date.day == 1 ) or found_expired_lease_daily or True:
            print("send mail")
            self.send_mail()
        return

def main():
    print("Lancement de l'application, veuillez patienter...")
    reminder = ReminderBot()
    reminder.apply()
       
if __name__ == "__main__":
    main()
    
