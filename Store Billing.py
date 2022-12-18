import pandas as pd
import smtplib
import email.message
import mimetypes 

from pathlib import Path

class Billings:

    def __init__(self):
        self.sales_df = pd.read_excel('Dataset/Vendas.xlsx')
        self.emails_df = pd.read_csv('Dataset/Emails.csv')
        self.stores_df = pd.read_csv('Dataset/Lojas.csv', sep=';', encoding='latin-1')

        self.sales_df = self.sales_df.merge(self.stores_df, on='ID Loja')
        self.last_date = self.sales_df['Data'].max()

        self.folders()
        self.send_emails()
        self.billings_ranking()
        self.board_email()

    def folders(self):
        # Saves the files in the folder.

        self.dict_stores = {}
        for store in self.stores_df['Loja']:
            self.dict_stores[store] = self.sales_df.loc[self.sales_df['Loja']==store, :]

        self.backup_path = Path('Stores Files Backup')
        self.folders_backup = self.backup_path.iterdir()
        names_list_backup = [folder.manager_name for folder in self.folders_backup]

        for store in self.dict_stores:
            # Checks if the folder does not exist and makes it.
            if store not in names_list_backup:
                Path(self.backup_path/store).mkdir()
            
            file_manager_name = f'{self.last_date.day}_{self.last_date.month}_{store}.csv'
            self.dict_stores[store].to_csv(self.backup_path/store/file_manager_name)

    def stores_result(self, store):
        #Target setting.
        self.target_billings_day = 1000
        self.target_billings_year = 1650000
        self.target_product_quantity_day = 4
        self.target_product_quantity_year = 120
        self.target_average_sales_day = 500
        self.target_average_sales_year = 500

            
        self.store_sales = self.dict_stores[store]
        self.store_sales_date = self.store_sales.loc[self.store_sales['Data']==self.last_date, :]

        # Store Billings.
        self.year_store_billings = self.store_sales['Valor Final'].sum()
        self.day_store_billings = self.store_sales_date['Valor Final'].sum()

        # How many different products were sold.
        self.product_quantity_year = len(self.store_sales['Produto'].unique())
        self.products_quantity_day = len(self.store_sales_date['Produto'].unique())

        # Average sale value on the year.
        self.sales_value_year = self.store_sales.groupby('Código Venda').sum(['Valor Final'])
        self.average_sales_year = self.sales_value_year['Valor Final'].mean()

        # Average sale value on the day.
        self.value_sales_day = self.store_sales_date.groupby('Código Venda').sum(['Valor Final'])
        self.average_sales_day = self.value_sales_day['Valor Final'].mean()  

    def send_emails(self):
      # Sending an e-mail to each manager with a spreadsheet
          for store in self.dict_stores:

            self.stores_result(store)

            if self.day_store_billings >= self.target_average_sales_day: color_bill_day = 'green' 
            else: color_bill_day = 'red'

            if self.year_store_billings >= self.target_billings_year: color_bill_year = 'green'
            else: color_bill_year = 'red'

            if self.products_quantity_day >= self.target_product_quantity_day: color_qtt_day = 'green'
            else: color_qtt_day = 'red'

            if self.product_quantity_year >= self.target_product_quantity_year: color_qtt_year = 'green'
            else: color_qtt_year = 'red'

            if self.average_sales_day >= self.target_average_sales_day: color_ticket_day = 'green'
            else: color_ticket_day = 'red'

            if self.average_sales_year >= self.target_average_sales_year: color_ticket_year = 'green'
            else: color_ticket_year = 'red'
            

            manager_name = self.emails_df.loc[self.emails_df['Loja']==store, 'Gerente'].values[0]
            body = f"""
                <p>Dear Manager, {manager_name}</p>

                <p>Yesterday's result <strong>({self.last_date.day}/{self.last_date.month})</strong> from the <strong>{store}</strong> store were:</p>

                <table>
                  <tr>
                    <th>Index</th>
                    <th>Day Value</th>
                    <th>Day Target</th>
                    <th>Day Scenario</th>
                  </tr>
                  <tr>
                    <td>Billings</td>
                    <td style="text-align: center">R${self.day_store_billings:.2f}</td>
                    <td style="text-align: center">R${self.target_billings_day:.2f}</td>
                    <td style="text-align: center"><font color="{color_bill_day}">◙</font></td>
                  </tr>
                  <tr>
                    <td>Product Diversity</td>
                    <td style="text-align: center">{self.products_quantity_day}</td>
                    <td style="text-align: center">{self.target_product_quantity_day}</td>
                    <td style="text-align: center"><font color="{color_qtt_day}">◙</font></td>
                  </tr>
                  <tr>
                    <td>Average Sales Value</td>
                    <td style="text-align: center">R${self.average_sales_day:.2f}</td>
                    <td style="text-align: center">R${self.target_average_sales_day:.2f}</td>
                    <td style="text-align: center"><font color="{color_ticket_day}">◙</font></td>
                  </tr>
                </table>
                <br>
                <table>
                  <tr>
                    <th>Index</th>
                    <th>Year Value</th>
                    <th>Year Target</th>
                    <th>Year Scenario</th>
                  </tr>
                  <tr>
                    <td>Billings</td>
                    <td style="text-align: center">R${self.year_store_billings:.2f}</td>
                    <td style="text-align: center">R${self.target_billings_year:.2f}</td>
                    <td style="text-align: center"><font color="{color_bill_year}">◙</font></td>
                  </tr>
                  <tr>
                    <td>Product Diversity</td>
                    <td style="text-align: center">{self.product_quantity_year}</td>
                    <td style="text-align: center">{self.target_product_quantity_year}</td>
                    <td style="text-align: center"><font color="{color_qtt_year}">◙</font></td>
                  </tr>
                  <tr>
                    <td>Average Sales Value</td>
                    <td style="text-align: center">R${self.average_sales_year:.2f}</td>
                    <td style="text-align: center">R${self.target_average_sales_year:.2f}</td>
                    <td style="text-align: center"><font color="{color_ticket_year}">◙</font></td>
                  </tr>
                </table>

                <p>Please find attached the spreadsheet with all the data for more details.</p>

                <p>Yours sincerely</p>
                <p>Francisco</p>
                """

            msg = email.message.EmailMessage()
            msg['Subject'] = f'OnePage of the Day {self.last_date.day}/{self.last_date.month} - Loja {store}'
            msg['From'] = '<e-mail>'
            msg['To'] = self.emails_df.loc[self.emails_df['Loja']==store, 'E-mail'].values[0]
            msg.add_header('Content-Type', 'text/html')
            msg.set_payload(body)
            password = '<password>'

            mime_type, _ = mimetypes.guess_type(f'{store}.csv')
            mime_type, mime_subtype = mime_type.split('/')

            file_path  = Path.cwd() / self.backup_path / store / f'{self.last_date.day}_{self.last_date.month}_{store}.csv'

            with open(file_path, 'rb') as file:
                msg.add_attachment(file.read(),
                maintype=mime_type,
                subtype=mime_subtype,
                filemanager_name=f'{store}.csv')

            server = smtplib.SMTP('smtp.gmail.com: 587')
            server.starttls()
            # Login Credentials for sending the mail
            server.login(msg['From'], password)
            server.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))


            print(f'Email to Store {store} sent!')

    def billings_ranking(self):
      # Billing ranking of all stores.
        total_billings = self.sales_df.groupby('Loja')[['Loja', 'Valor Final']].sum('Valor Final')
        self.annual_billings = total_billings.sort_values(by='Valor Final', ascending=False)

        annual_rk_file = f'{self.last_date.day}_{self.last_date.month}_Annual Ranking.csv'
        self.annual_billings.to_csv(f'Stores Files Backup/{annual_rk_file}')

        self.store_sales_date = self.sales_df.loc[self.sales_df['Data']==self.last_date, :]
        self.stores_billings_day = self.store_sales_date.groupby('Loja')[['Loja', 'Valor Final']].sum('Valor Final')
        self.stores_billings_day = self.stores_billings_day.sort_values(by='Valor Final', ascending=False)

        days_rk_file = f"{self.last_date.day}_{self.last_date.month}_Days Ranking.csv"
        self.stores_billings_day.to_csv(f'Stores Files Backup/{days_rk_file}')

        
    def board_email(self):
      # Sending an e-mail to board of directors.
        year_best_store = f'R${self.annual_billings.iloc[0, 0]:,.2f}'.replace(',', '_').replace('.', ',').replace('_', '.')
        year_worse_store = f'R${self.annual_billings.iloc[-1, 0]:,.2f}'.replace(',', '_').replace('.', ',').replace('_', '.')

        day_best_store = f'R${self.stores_billings_day.iloc[0, 0]:,.2f}'.replace(',', '_').replace('.', ',').replace('_', '.')
        day_worse_store = f'R${self.stores_billings_day.iloc[-1, 0]:,.2f}'.replace(',', '_').replace('.', ',').replace('_', '.')

        body = f"""
        <p>To whom it may concern,</p>
        
        <p>Best Store of the Day in Billing: <strong>Loja {self.stores_billings_day.index[0]}</strong> with revenues of <strong>{day_best_store}</strong></p>
        <p>Worst Store of the Day in Billing: <strong>Loja {self.stores_billings_day.index[-1]}</strong> with revenues of <strong>{day_worse_store}</strong></p>
        <br>
        <p>Best Store of the Year in Billing: <strong>Loja {self.annual_billings.index[0]}</strong> with revenues of <strong>{year_best_store}</strong></p>
        <p>Worst Store of the Year in Billing: <strong>Loja {self.annual_billings.index[-1]}</strong> with revenues of <strong>{year_worse_store}</strong></p>
        <br>
        <p>Please find attached the year and day ratings of all stores.</p>
        <br>
        <p>Yours sincerely</p>
        <p>Francisco</p>
        """

        msg = email.message.EmailMessage()
        msg['Subject'] = f'Store Ratings {self.last_date.day}/{self.last_date.month}'
        msg['From'] = '<e-mail>'
        msg['To'] = self.emails_df.loc[self.emails_df['Loja']=='Diretoria', 'E-mail'].values[0]
        msg.add_header('Content-Type', 'text/html')
        msg.set_payload(body)
        password = '<password>' 


        mime_type, _ = mimetypes.guess_type('Days Ranking.csv')
        mime_type, mime_subtype = mime_type.split('/')

        annual_rk_path  = Path.cwd() / self.backup_path / f'{self.last_date.day}_{self.last_date.month}_Annual Ranking.csv'
        days_rk_path  = Path.cwd() / self.backup_path / f'{self.last_date.day}_{self.last_date.month}_Days Ranking.csv'


        with open(annual_rk_path, 'rb') as file:
            msg.add_attachment(file.read(),
            maintype=mime_type,
            subtype=mime_subtype,
            filemanager_name='Annual Ranking.csv')

        with open(days_rk_path, 'rb') as file:
            msg.add_attachment(file.read(),
            maintype=mime_type,
            subtype=mime_subtype,
            filemanager_name='Days Ranking.csv')


        server = smtplib.SMTP('smtp.gmail.com: 587')
        server.starttls()
        # Login Credentials for sending the mail
        server.login(msg['From'], password)
        server.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))


        print(f'E-mail to board of directors sent!')



if __name__ == '__main__':
    store_billings = Billings()