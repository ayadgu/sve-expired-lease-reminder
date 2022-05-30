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

        self.found_expired_lease_monthly=False
        self.found_expired_lease_daily=False
        

    def initDataFrame(self):
        pd.set_option('display.max_rows', None)
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', None)

        #--------------------------
        # DF Fichier Code
                
        filename=glob.glob("*baux*.xlsx")[0]
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
        # print(tabulate(df_fichier_baux, headers='keys', tablefmt='psql'))
        # print(df_fichier_baux[col_nom_locataire])

        df_new_df=pd.concat([df_fichier_baux[col_numero_bail],df_fichier_baux[col_nom_locataire],df_fichier_baux[col_type_bail],df_fichier_baux[col_fin_bail]], axis=1)
        print(df_new_df.head(5))
        
        self.df_new_df = df_new_df

    def send_mail(self):
        text = f"""
        <h6>Ceci est un message automatique.</h6>

        <h2>Baux arrivés à expiration:</h2>

        {self.expired_bails_monthly.get_html_string()}
        {self.expired_bails_daily.get_html_string()}

        Cordialement
        """

        html = f"""
        <html><body>
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

        <table style="display:none;">
            <tr>
                <th>Code</th>
                <th>Signification</th>
            </tr>
            <tr>
                <td>0</td>
                <td>Code Civil ICC</td>
            </tr>
            <tr>
                <td>1</td>
                <td>Loi Quillot 3 ans</td>
            </tr>
            <tr>
                <td>2</td>
                <td>Loi Quillot 6 ans</td>
            </tr>
            <tr>
                <td>3</td>
                <td>Meublé</td>
            </tr>
            <tr>
                <td>4</td>
                <td>Commercial</td>
            </tr>
            <tr>
                <td>5</td>
                <td>Professionnel</td>
            </tr>
            <tr>
                <td>6</td>
                <td>Loi de 48</td>
            </tr>
            <tr>
                <td>7</td>
                <td>Bail de 6 ans</td>
            </tr>
            <tr>
                <td>8</td>
                <td>Bail Mehaignerie 3 ans</td>
            </tr>
            <tr>
                <td>9</td>
                <td>Bail Mehaignerie 8 ans</td>
            </tr>
            <tr>
                <td>A</td>
                <td>Bail Mehaignerie 6 ans</td>
            </tr>
            <tr>
                <td>B</td>
                <td>Bail Loi 06.07.89 3 ans</td>
            </tr>
            <tr>
                <td>C</td>
                <td>Bail Loi 06.07.89 6 ans</td>
            </tr>
            <tr>
                <td>D</td>
                <td>Bail Loi 07.89 3 ans 1/6</td>
            </tr>
            <tr>
                <td>E</td>
                <td>Bail comm. Dg Non Revisable</td>
            </tr>
            <tr>
                <td>F</td>
                <td>Bail SRU</td>
            </tr>
            <tr>
                <td>G</td>
                <td>Bail Derog. (art L.145-5 cc)</td>
            </tr>
            <tr>
                <td>H</td>
                <td>Code Civil IRL</td>
            </tr>
        </table>

        </p>
        <p>Bonne journée !</p>
        </body></html>
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
        expired_bails_monthly = PrettyTable()
        expired_bails_daily = PrettyTable()

        expired_bails_monthly.field_names  = ["Numéro Mandat","Nom Locataire","Type Bail","Date Fin"]
        expired_bails_daily.field_names  = ["Numéro Mandat","Nom Locataire","Type Bail","Date Fin"]

        for indices, row in self.df_new_df.iterrows():
            date = self.df_new_df.at[indices,"Fin Bail"]
            date = datetime.strptime(date, '%d/%m/%Y')
            print("")
            numero_bail = self.df_new_df.at[indices,"Numero Mandat"]
            nom_locataire = self.df_new_df.at[indices,"Nom Locataire"]
            type_bail = str(self.df_new_df.at[indices,"Type Bail"])
            empty_row=["Aucun"]*len(self.df_new_df.columns)

            match type_bail :
                # #########
                # Code bails
                # # 4 : Prévenir à la date de fin et (date de fin + 3 ans - 6 mois)
                # 0, 3, B, C, G : Prévenir 8 mois avant date de fin
                
                case '4':
                    if self.today_date == date or self.today_date > date + self.three_years-self.six_months:
                        self.found_expired_lease_monthly=True
                        expired_bails_monthly.add_row([numero_bail,nom_locataire,type_bail,date.strftime('%d/%m/%Y')])
                    
                    if self.today_date == date or self.today_date == date + self.three_years-self.six_months:
                        self.found_expired_lease_daily=True
                        expired_bails_daily.add_row([numero_bail,nom_locataire,type_bail,date.strftime('%d/%m/%Y')])
                
                case _:
                    if self.today_date > date-self.height_months:
                        self.found_expired_lease_monthly=True
                        expired_bails_monthly.add_row([numero_bail,nom_locataire,type_bail,date.strftime('%d/%m/%Y')])
                        
                    if self.today_date == date-self.height_months:
                        self.found_expired_lease_daily=True
                        expired_bails_daily.add_row([numero_bail,nom_locataire,type_bail,date.strftime('%d/%m/%Y')])

        self.expired_bails_monthly=expired_bails_monthly
        self.expired_bails_daily=expired_bails_daily

        print("expired_bails_daily : \n",self.expired_bails_daily)
        print("expired_bails_monthly : \n",self.expired_bails_monthly)
        print(self.today_date.day == 1, "first day")

        print("self.found_expired_lease_monthly",self.found_expired_lease_monthly)
        print("self.found_expired_lease_daily",self.found_expired_lease_daily)

        if not self.found_expired_lease_daily :
            self.expired_bails_daily.add_row(empty_row)

        if not self.found_expired_lease_monthly :
            self.expired_bails_monthly.add_row(empty_row)

        print(self.date_fichier_excel)

        if (self.found_expired_lease_monthly and not self.today_date.day == 1 ) or self.found_expired_lease_daily:
            print("send mail")
            self.send_mail()
        return

def main():
    print("Lancement de l'application, veuillez patienter...")
    reminder = ReminderBot()
    reminder.apply()
       
if __name__ == "__main__":
    main()
    
