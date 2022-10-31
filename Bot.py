from botcity.core import DesktopBot
import pandas as pd
import win32com.client as win32

pagamento = pd.read_excel('PENDENCIAS.xlsm')


class Bot(DesktopBot):
    def action(self, execution=None):



        self.execute(r'C:\Users\Otavio.cunha\Desktop\Protheus 33.lnk')

        if not self.find( "OK", matching=0.97, waiting_time=90000):
            self.not_found("OK")
        self.click()
        if not self.find( "BemVindo", matching=0.97, waiting_time=10000):
            self.not_found("BemVindo")
        self.paste('')
        self.tab()
        self.paste('')
        if not self.find( "Entrarr", matching=0.97, waiting_time=10000):
            self.not_found("Entrarr")
        self.click()
        self.wait(10000)
        self.tab()
        self.tab()
        self.paste('00601000')
        self.paste('06')
        self.tab()
        self.enter()
        self.wait(8000)
        if not self.find( "Favoritos", matching=0.97, waiting_time=90000):
            self.not_found("Favoritos")
        self.click()
        if not self.find( "Funções", matching=0.97, waiting_time=90000):
            self.not_found("Funções")
        self.click()
        self.wait(10000)
        self.paste('12092022')
        if not self.find( "confi", matching=0.97, waiting_time=90000):
            self.not_found("confi")
        self.click()
        if not self.find( "ctas a receber", matching=0.97, waiting_time=90000):
            self.not_found("ctas a receber")
        self.click()
        if not self.find( "incluir", matching=0.97, waiting_time=90000):
            self.not_found("incluir")
        self.click()
        self.wait(10000)
        self.tab()
        self.tab()
        self.paste('00601000')
        self.enter()
        self.enter()
        self.enter()

        for linha in pagamento.index:

            valor = pagamento.loc[linha, 'Valor']
            cliente = pagamento.loc[linha, 'Cliente']
            Data = pagamento.loc[linha, 'Data']
            Loja = pagamento.loc[linha,'Loja']
            historico = pagamento.loc[linha,'Historico']

            if not self.find( "Dados", matching=0.97, waiting_time=90000):
                self.not_found("Dados")
            self.wait(25000)
            self.paste('MAN')
            self.tab()
            self.tab()
            self.paste('RA')
            self.tab()
            if not self.find( "Entrada", matching=0.97, waiting_time=90000):
                self.not_found("Entrada")
            self.wait(8000)
            self.paste('033')
            self.paste('2271')
            self.tab()
            self.wait(8000)
            self.paste('130105228')
            if not self.find( "OKAY", matching=0.97, waiting_time=90000):
                self.not_found("OKAY")
            self.click()
            self.click()
            self.wait(8000)
            self.paste('1305')
            self.tab()
            self.paste(f'{cliente}'.replace(".0","").zfill(6))
            self.wait(5000)
            self.paste(f'{Loja}'.replace(".0","").zfill(2))
            self.tab()
            self.paste(Data)
            self.tab()
            self.paste(f'{valor:.2f}'.replace(".",","))
            self.wait(5000)
            if historico == 0:
                pass
            else:
                self.paste(historico)
            if not self.find( "Salvar", matching=0.97, waiting_time=90000):
                self.not_found("Salvar")
            self.click()
            self.wait(10000)
        if not self.find( "EXCLUIR", matching=0.97, waiting_time=10000):
            self.not_found("EXCLUIR")
        self.click()
        if not self.find( "FINISh", matching=0.97, waiting_time=10000):
            self.not_found("FINISh")
        self.click()



    def not_found(self, label):
        print(f"Element not found: {label}")




if __name__ == '__main__':
    Bot.main()


print('e-mail enviado')
