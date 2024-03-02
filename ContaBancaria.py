from openpyxl import Workbook
from datetime import datetime


class ContaBancaria:
    def __init__ (self):
        # Inicializa os atributos da conta bancária
        self.saldo = 0
        self.depositos = []  # Lista para armazenar os valores depositados
        self.saques = []     # Lista para armazenar os valores sacados
        self.qtdsaquediario = 3  # Limite de saques diários
        self.qtdsaquesfeitos = 0  # Contador para controlar os saques diários
        
    def depositar(self, valor):
        # Método para realizar um depósito
        if valor > 0:
            self.saldo += valor
            self.depositos.append(valor)
            print(f"\nO Depósito do valor R$ {valor} foi realizado com sucesso!")
            
    def sacar(self, valor): 
        # Método para realizar um saque
        if self.qtdsaquesfeitos < self.qtdsaquediario:
            if self.saldo >= valor:
                self.saldo -= valor
                self.saques.append(valor)
                self.qtdsaquesfeitos += 1
                print(f"\nO Saque do valor R$ {valor} foi realizado com sucesso!")
            else:
                print("\nValor de Saldo insuficiente")
        else:
            print("\nLimite de saque diário atingido.\nTente novamente amanhã.")
    
    def exportarparaexcel(self, filename):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Extrato Bancário"
        
        sheet['A1'] = "Data e Hora da Exportação"
        sheet['B1'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        sheet['A3'] = "Saldo Atual da Conta"
        sheet['B3'] = self.saldo
        
        sheet['A5'] = "Depósitos"
        sheet['A6'] = "ID"
        sheet['B6'] = "Valor"
        for idx, deposito in enumerate(self.depositos, start=1):
            sheet[f'A{idx+6}'] = idx
            sheet[f'B{idx+6}'] = deposito
            
        sheet['C5'] = ""  # Coluna vazia entre depósitos e saques
        
        sheet['D5'] = "Saques"
        sheet['D6'] = "ID"
        sheet['E6'] = "Valor"
        for idx, saque in enumerate(self.saques, start=1):
            sheet[f'D{idx+6}'] = idx
            sheet[f'E{idx+6}'] = saque
            
        workbook.save(filename)
        print(f"Extrato exportado para {filename} com sucesso!")
        
    def extrato(self):
        # Método para exibir o extrato da conta bancária
        print("----------------------------------------\n EXTRATO\n----------------------------------------")
        print(f"Saldo Atual da Conta: {self.saldo}\n")
        print("Histórico:")
        print("----------------------------------------")
        print("Movimentação de Depósitos da Conta:")
        print("ID    | Depósito")
        print("----------------------------------------")
        
        for id, deposito in enumerate(self.depositos, start=1):
            print(f"{id:<5} | R$ {deposito}")
            
        print("----------------------------------------")
        print("Movimentação de Saques da Conta:")
        print("ID    | Saque")
        print("----------------------------------------")
        
        for id, saque in enumerate(self.saques, start=1):
            print(f"{id:<5} | R$ {saque}")
                
        print("----------------------------------------")

# Função principal
if __name__ == "__main__":
    i = True
    minhaConta = ContaBancaria()

    while i:
        #Menu de opções
        print("Escolha uma operação:")
        print("1- Depositar")
        print("2- Sacar")
        print("3- Extrato")
        print("4- Exportar informações do Extrato para Excel")
        print("5- Sair")
        
        op = input("Digite o número do Menu da operação desejada: ")
        
        if op == "1":
            entrada = float(input("Digite o valor do depósito: "))
            minhaConta.depositar(entrada)
            
        elif op == "2":
            saida = float(input("Digite o valor que será sacado do saldo da conta: "))
            minhaConta.sacar(saida)
        
        elif op == "3":
            print("Será exibido o saldo da conta acompanhado com suas movimentações separadas entre depósito e saque.\n")
            minhaConta.extrato()
            
        elif op == "4":
            minhaConta.exportarparaexcel("extratobancario.xlsx")
            print("Informações do Saldo exportado com sucesso para planilha de Excel!")
    
        elif op == "5":
            print("Encerrando operações bancárias!\n")
            i = False
        
        else:
            print("\nOpção inválida, tente novamente.\n")