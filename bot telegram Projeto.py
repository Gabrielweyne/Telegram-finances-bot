import telebot
import pandas as pd
import datetime


bot = telebot.TeleBot('PRIVATE TOKEN')


@bot.message_handler(commands=['start', 'menu'])
def responder(mensagem):
    texto = '''Olá, esse é o menu de opções, o que você deseja fazer?
/opcao0 para caso nunca tenha usado o bot
/opcao1 para mostrar saldo
/opcao2 para mostrar gastos que teve no mês
/opcao3 para adicionar saldo
/opcao4 para registrar uma despesa
/opcao5 para analise das despesas e histórico
/opcao6 para ver a sua planilha

'''
    bot.reply_to(mensagem, texto)

@bot.message_handler(commands=['opcao0'])
def opcao0(mensagem):
    boas_vindas='''Seja bem vindo(a) ao bot de finanças.\nAqui você adicionará seu saldo, quanto você gasta de despesa, e também a área do investimento.Você gostaria de adicionar algum saldo inicial ?'''
    mensagem=bot.send_message(mensagem.chat.id, 'Seja bem vindo(a) ao bot finanças. Informe seu email para criarmos um cadastro para você.')
    bot.register_next_step_handler(mensagem, cadastro)




@bot.message_handler(commands=['opcao1'])
def opcao1(mensagem):
    #Opção 1 está errada! Usasse o dicionario finanças, mas ele não é global.
    nome=mensagem.from_user.first_name
    planilha_df=pd.read_excel(f"{nome}.xlsx")
    numero_do_mês = datetime.datetime.fromtimestamp(mensagem.date).month
#temos que descobrir o mês.
    meses=["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
    mês=meses[numero_do_mês-1]

    
    saldo=planilha_df.loc[planilha_df[" "] == "Saldo final", f"{mês}"].iloc[0]
    bot.send_message(mensagem.chat.id, f'Esse é o seu saldo atual: R$ {saldo:.2f}')

@bot.message_handler(commands=['opcao2'])
def opcao2(mensagem):
    #ler na planilha e mostrar. (CHECK)
    nome=mensagem.from_user.first_name
    planilha_df=pd.read_excel(f"{nome}.xlsx")
    numero_do_mês=datetime.datetime.fromtimestamp(mensagem.date).month
    meses=["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
    mês=meses[numero_do_mês-1]
    planilha_df[[f'{mês}']]
#a opção 2 tá errada, precisamos ler a planilha
    for i in range(0,len(finanças)):
        
        x=planilha_df[['finanças[mês][' '][i]']]
        texto1=f"Você gastou {x} com {finanças[' '][i]}"
        texto2=f"No mês {mês} você gastou com:"
        texto2=texto2+\ntexto1
    
    bot.send_message(mensagem.chat.id, f"{texto2}")

@bot.message_handler(commands=['opcao3'])
def opcao3(mensagem):
    #CHECK OBSERVAÇÃO DE ERRO GROTESCO, TIRAR BOT SALDO=MENSAGEM.TEXT
    #digitando o saldo, a gente coloca na planilha
    mensagem=bot.send_message(mensagem.chat.id, "Digite quanto você deseja adicionar ao saldo:")
    #'''observação, essa função bot.register_next_step_handler serve para receber a mensagem e gerar uma função. Portanto, para tudo precisamos criar uma função.'''
    bot.register_next_step_handler(mensagem,processar_saldo)
def processar_saldo(mensagem):
    try:
        valor=float(mensagem.text.replace(',','.'))
        if valor<=0:
            bot.send_message(mensagem.chat.id, "Valor inválido!")
        else:
            numero_do_mês=datetime.datetime.fromtimestamp(mensagem.date).month
            meses=["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
            mês=meses[numero_do_mês-1]
            nome=mensagem.from_user.first_name
            planilha_df = pd.read_excel(f"{nome}.xlsx")
            planilha_df.loc[planilha_df[" "] =="Saldo depositado",mês]=planilha_df.loc[planilha_df[" "] =="Saldo depositado",mês]+valor
            planilha_df.loc[planilha_df[" "] =="Saldo final",mês]=planilha_df.loc[planilha_df[" "] =="Saldo final",mês]+valor
            planilha_df.to_excel(f"{nome}.xlsx", index=False)
            df=pd.read_excel(f"{nome}.xlsx")
            x=planilha_df.loc[planilha_df[" "] == "Saldo final", mês].values[0]
            bot.send_message(mensagem.chat.id, f"Esse é o seu saldo atual, após o deposito: {x}")
                
    except FileNotFoundError:
        bot.send_message(mensagem.chat.id, "Você ainda não tem cadastro. Use /opcao0 primeiro.")
    

@bot.message_handler(commands=['opcao4'])
def opcao4(mensagem):
    texto='''Digite qual foi seu gasto:
/Saude
/Comida
/Uber
/Gasolina
/Academia
/Lazer
/Roupa
/Streaming
/Igreja
/Despesa_adicional
'''
    mensagem=bot.send_message(mensagem.chat.id, texto)
    bot.register_next_step_handler(mensagem,despesas)

    #O codigo está errado aqui, vamos fazer diferente ! Vamos ler a planilha pelo modo read do pandas.(CHECK)
def despesas(mensagem):
    gasto = mensagem.text.replace('/','')

    nome=mensagem.from_user.first_name
    planilha_df=pd.read_excel(f"{nome}.xlsx")
    numero_do_mês=datetime.datetime.fromtimestamp(mensagem.date).month
    meses=["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
    mês=meses[numero_do_mês-1]
    
    bot.send_message(mensagem.chat.id,f'Qual foi a sua despesa com {gasto}')
    valor_gasto = mensagem.text.replace('')

    planilha_df.loc[planilha_df[" "] == f"{mês}", f"{gasto}"]=planilha_df.loc[planilha_df[" "] == f"{mês}", f"{gasto}"]- valor_gasto

    gasto = mensagem.text.replace('/', '')
    bot.send_message(mensagem.chat.id,f'Qual foi a sua despesa com {gasto}')
    #'''Agora, vamos fazer o seguinte: Ler a planilha do cara, e alterar os dados de gastos para qual ele gastou.'''
    posicao=finanças[" "].index(gasto)
    mes=datetime.datetime.fromtimestamp(mensagem.date).month
    #'''Vamos fazer um if e else para garantir que conseguimos abrir ou não a planilha da pessoa'''

    


@bot.message_handler(commands=['opcao5'])
def opcao5(mensagem):
    texto='''
/opcao51 Saber despesa de meses passados
/opcao52 Comparar despesas de mes em mes
/opcao53 Para ter acesso ao seu histórico

'''
    bot.send_message(mensagem.chat.id, texto)

    
@bot.message_handler(commands=['opcao51'])
def opcao51(mensagem):
    texto='''Digite qual foi seu gasto:
/Saude
/Comida
/Uber
/Gasolina
/Academia
/Lazer
/Roupa
/Streaming
/Igreja
/Despesa_adicional
'''
    mensagem=bot.send_message(mensagem.chat.id, texto)
    bot.register_next_step_handler(mensagem,descobrindo_gasto)
def descobrindo_gasto(mensagem):
    gasto = mensagem.text.replace('/', '')
    
    
    texto=f'''
Aperte em qual mês você quer saber o seu gasto com {gasto}
/Janeiro
/Fevereiro
/Março
/Abril
/Maio
/Junho
/Julho
/Agosto
/Setembro
/Outubro
/Novembro
/Dezembro
'''
    mensagem=bot.send_message(mensagem.chat.id, texto)
    bot.register_next_step_handler(mensagem,mostrando_gasto)
    def mostrando_gasto(mensagem):
        mês=mensagem.text.replace('/', '')
        '''Vamos fazer aqui o seguinte: Abrir a planilha no modo read do excel, acharmos printamos na conversa.'''
        nome=mensagem.from_user.first_name
        for i in range(0,len(finanças[" "])):
            if gasto==finanças[" "][i]:
                posição_gasto=i
        for i in range(0,len(finanças)):
            if mês==finanças[i]:
                posição_mês=i
            
        planilha_df=pd.read_excel(f'{nome}.xlsx')
    
        print(planilha_df[[posição_gasto][posição_mês]])
        bot.send_message(mensagem.chat.id,planilha_df[[f'{gasto}'][f'{mês}']] )
    


'''@bot.message_handler(commands=['opcao6'])
def opcao6(mensagem):
    #enviar planilha'''







def cadastro(mensagem):
    email=mensagem.text
    bot.send_message(mensagem.chat.id, f"Seu email é {email}")
    ''' saber nome da pessoa por id'''
    finanças={
        " ":["Saldo depositado","Saldo final","Academia","Uber","Saúde","Streaming","Igreja","Roupa","Lazer","Observações"],
        "Janeiro":[0,0,0,0,0,0,0,0,0,0],
        "Fevereiro":[0,0,0,0,0,0,0,0,0,0],
        "Março":[0,0,0,0,0,0,0,0,0,0],
        "Abril":[0,0,0,0,0,0,0,0,0,0],
        "Maio":[0,0,0,0,0,0,0,0,0,0],
        "Junho":[0,0,0,0,0,0,0,0,0,0],
        "Julho":[0,0,0,0,0,0,0,0,0,0],
        "Agosto":[0,0,0,0,0,0,0,0,0,0],
        "Setembro":[0,0,0,0,0,0,0,0,0,0],
        "Outubro":[0,0,0,0,0,0,0,0,0,0],
        "Novembro":[0,0,0,0,0,0,0,0,0,0],
        "Dezembro":[0,0,0,0,0,0,0,0,0,0],
        }
    nome=mensagem.from_user.first_name
    planilha_df=pd.DataFrame(finanças)
    planilha_df.to_excel(f"{nome}.xlsx",index=False)

        
bot.polling()

