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
from dateutil.relativedelta import relativedelta

class ReminderBot():
    def __init__(self):
        self.initVariable()
        self.initDataFrame()

    def initVariable(self):
        load_dotenv()
        os.path.expanduser('~')

        self.six_months = timedelta(6*365/12)
        self.eight_months = timedelta(8*365/12)
        self.three_years = timedelta(3*365)
        self.today_date = datetime.today().date()


        self.PASSWORD_OUTLOOK=os.environ.get("PASSWORD_OUTLOOK")
        self.EMAIL_OUTLOOK=os.environ.get("EMAIL_OUTLOOK")
        self.EMAIL_DEST=os.environ.get("EMAIL_DEST")

        

    def initDataFrame(self):
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)

        # A surveiller car mauvaise practice
        # Permet de supress le warning A value is trying to be set on a copy of a slice from a DataFrame
        pd.options.mode.chained_assignment = None

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
        
        self.col_fin_bail=col_fin_bail 
        self.col_type_bail=col_type_bail 
        self.col_numero_bail=col_numero_bail 
        self.col_nom_locataire=col_nom_locataire 

        df_fichier_baux=df_fichier_baux.rename(columns={19: col_fin_bail})
        df_fichier_baux=df_fichier_baux.rename(columns={17: col_type_bail})
        df_fichier_baux=df_fichier_baux.rename(columns={0: col_numero_bail})
        df_fichier_baux=df_fichier_baux.rename(columns={3: col_nom_locataire})

        df_fichier_baux[col_type_bail] = df_fichier_baux[col_type_bail].astype(str)


        df_fichier_baux[col_nom_locataire].bfill(inplace=True)

        df_fichier_baux = df_fichier_baux[df_fichier_baux[col_fin_bail].notna()]
        df_fichier_baux = df_fichier_baux[df_fichier_baux[col_numero_bail].notna()]
        
        df_situation_des_baux=df_fichier_baux[[col_numero_bail,col_nom_locataire,col_type_bail,col_fin_bail]]

        df_situation_des_baux[col_fin_bail] = pd.to_datetime(df_situation_des_baux[col_fin_bail],format= '%d/%m/%Y').dt.date



        df_situation_des_baux.sort_values(by=col_fin_bail,inplace=True,ascending=False)


        self.df_situation_des_baux = df_situation_des_baux

        inner_join = pd.merge(df_situation_des_baux, 
                                            df_adresse_baux_min, 
                                            on =col_numero_bail, 
                                            how ='inner')
        inner_join=inner_join.loc[:, [col_numero_bail,col_type_bail,col_nom_locataire,col_adresse_bail,col_fin_bail]]

        inner_join=self.delete_unwanted_rows(inner_join,col_type_bail,col_fin_bail)

        inner_join[col_type_bail]=self.format_type_bail(inner_join,col_type_bail)
        self.inner_join=inner_join



    def format_date_us_to_eur(self,df,col):
        df[col] = pd.to_datetime(df[col], errors='coerce')
        return df[col].dt.strftime('%d-%m-%Y')

    def format_type_bail(self,df,col):
        return df[col].replace(["0","1","2","3","4","5","6","7","8","9","A","B","C","D","E","F","G","H"],["0 Code Civil ICC","1 Loi Quillot 3 ans","2 Loi Quillot 6 ans","3 Meublé","4 Commercial","5 Professionnel","6 Loi de 48","7 Bail de 6 ans","8 Bail Mehaignerie 3 ans","9 Bail Mehaignerie 8 ans","A Bail Mehaignerie 6 ans","B Bail Loi 06.07.89 3 ans","C Bail Loi 06.07.89 6 ans","D Bail Loi 07.89 3 ans 1/6","E Bail comm. Dg Non Revisable","F Bail SRU","G Bail Derog. (art L.145-5 cc)","H Code Civil IRL"])
        
    def delete_unwanted_rows(self,df,col_type_bail,col_fin_bail):
        df.drop(df[(df[col_type_bail] == "4") & (self.today_date < df[col_fin_bail] + self.three_years-self.six_months) ].index, inplace=True)
        df.drop(df[(df[col_type_bail] != "4") & (self.today_date < df[col_fin_bail]-self.eight_months)].index, inplace=True)
        return df

    def send_mail(self):
        text = f"""
        <h6>Ceci est un message automatique.</h6>

        <h2>Baux arrivés à expiration:</h2>
        {self.inner_join.to_html(index=False,classes=["table-bordered", "table-striped", "table-hover"])}

        Cordialement
        """

        html = f"""
        <html>
              
        <html>
            <body>
                <p>Ceci est un message automatique.</p>
                <h3>Baux bientôt expirés depuis aujourd'hui exactement :</h3>
                {self.expired_bails_daily.to_html(index=False)}
                <br/>
                <h3>Baux bientôt ou déjà expirés à ce jour :</h3>
                {self.inner_join.to_html(index=False)}
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
        # msg.set_content((self.expired_bails_daily.to_html(index=False)))
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
        self.today_date+= relativedelta(months=6)

        # Init PrettyTable
        today_expired_df=self.inner_join[\
            ((self.inner_join["Type Bail"]=="4") &\
                (\
                    (self.today_date==self.inner_join["Fin Bail"])|\
                    ((self.today_date==self.inner_join["Fin Bail"] + self.three_years-self.six_months))\
                )\
            )|\
            ((self.inner_join["Type Bail"]!="4") &\
            (\
                (self.today_date==self.inner_join["Fin Bail"])|\
                (self.today_date==self.inner_join["Fin Bail"]-self.eight_months)\
            )\
            )
        ]
        
        if today_expired_df:
            print("found_expired_lease_daily",True)
        else:
            print("found_expired_lease_daily",False)

        if not today_expired_df :
            empty_row=["Aucun"]*len(self.inner_join.columns)
            today_expired_df.append(empty_row)
        
        self.inner_join["Fin Bail"]=self.format_date_us_to_eur(self.inner_join,"Fin Bail")

        print("today_expired_df : \n",today_expired_df)
        self.today_expired_df=today_expired_df

        if (self.today_date in self.inner_join[self.col_fin_bail].values)\
            or self.today_date.day == 1\
            or today_expired_df\
            or 1 :
            print("email sent")
            self.send_mail()

def main():
    print("Lancement de l'application, veuillez patienter...")
    reminder = ReminderBot()
    reminder.apply()
       
if __name__ == "__main__":
    main()
    
