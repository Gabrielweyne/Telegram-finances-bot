import telebot
import pandas as pd
import datetime

bot = telebot.TeleBot('PRIVATE TOKEN')


@bot.message_handler(commands=['start', 'menu'])
def responder(mensagem):
    texto = '''Olá, esse é o menu de opções, o que você deseja fazer?
/opcao0 para caso nunca tenha usado o bot
/opcao1 para mostrar saldo
/opcao2 para mostrar gastos que teve no mês ou meses anteriores
/opcao3 para adicionar saldo
/opcao4 para registrar uma despesa
/opcao6 para ver a sua planilha
/opcao7 corrigir um valor enviado errado
/opcao8 fazer uma analise dos gastos mensais
/opcao9 área de investimentos
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
    def mostrar_gastos(mensagem):
        texto=''
        mês= mensagem.text.replace('/', '')
        bot.send_message(mensagem.chat.id, "Esses foram seus gastos:")
        for i in range(len(planilha_df[" "])):
            gasto=planilha_df[" "][i]
            valor_do_gasto=planilha_df.loc[planilha_df[" "]==gasto,mês].values[0]
            texto=texto+ f"{gasto}: R$ {valor_do_gasto:.2f}\n"
        bot.send_message(mensagem.chat.id,texto)
    #ler na planilha e mostrar. (CHECK)
    nome=mensagem.from_user.first_name
    planilha_df=pd.read_excel(f"{nome}.xlsx")
    numero_do_mês=datetime.datetime.fromtimestamp(mensagem.date).month
    mensagem=bot.send_message(mensagem.chat.id, "De qual mês você quer saber os gastos ? \n/Janeiro \n/Fevereiro \n/Março \n/Abril \n/Maio \n/Junho \n/Julho \n/Agosto \n/Setembro \n/Outubro \n/Novembro \n/Dezembro")
    bot.register_next_step_handler(mensagem,mostrar_gastos)

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
/Saúde
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



    
    def valor_despesa(mensagem):
        valor_gasto = mensagem.text.replace('R$','').replace('-', '').strip()
        valor_gasto=int(valor_gasto)
    
        planilha_df.loc[planilha_df[" "] == f"{gasto}", f"{mês}"]=planilha_df.loc[planilha_df[" "] == f"{gasto}", f"{mês}"]- valor_gasto

        planilha_df.loc[planilha_df[" "] == "Saldo final", f"{mês}"]=planilha_df.loc[planilha_df[" "] == "Saldo final", f"{mês}"]- valor_gasto
        '''Essa linha aqui em cima é para o Saldo Final'''
        Saldo_final=planilha_df.loc[planilha_df[" "] == "Saldo final", f"{mês}"].values[0]
        bot.send_message(mensagem.chat.id,f'Computamos na planilha seu gasto com {gasto}')
        bot.send_message(mensagem.chat.id,f'Seu saldo atual é {Saldo_final}')
        planilha_df.to_excel(f"{nome}.xlsx", index=False)
    

    mensagem=bot.send_message(mensagem.chat.id,f'Qual foi a sua despesa com {gasto}. Favor digite o valor positivo \nExemplo: R$1000')
    bot.register_next_step_handler(mensagem,valor_despesa)

    



'''@bot.message_handler(commands=['opcao6'])
def opcao6(mensagem):
    #enviar planilha'''

@bot.message_handler(commands=['opcao7'])#a opção está muito errada, não sei corrigir ela, vou estudar cálculo.
def corrigindo_erro(mensagem):
    def descobrir_erro(mensagem,mês):
            gasto=mensagem.text.replace('/','')
            if (gasto=='Saldo_depositado'):
                gasto='Saldo depositado'
                nome=mensagem.from_user.first_name
                planilha_df=pd.read_excel(f"{nome}.xlsx")
                saldo_antigo=planilha_df.loc[planilha_df[' ']=="Saldo depositado",f"{mês}"].values[0]
                mensagem=bot.send_message(mensagem.chat.id,f"Esse é saldo que estava antes de você querer alterar. Você tem certeza que quer trocar ?\n/Sim.\n/Não.")
                bot.register_next_step_handler(mensagem,validar_resposta)
                valor_para_trocar=bot.register_next_step_handler(mensagem,valor_correto)
                planilha_df.loc[planilha_df[' ']==["Saldo depositado",f"{mês}"]]=valor_para_trocar
                planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]=valor_para_trocar
                '''Agora, vamos corrigir, pegar os valores de despesa e tirar do valor de Saldo final.'''
                planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]=planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]-planilha_df.loc[planilha_df[' ']==["Academia",f"{mês}"]]
                planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]=planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]-planilha_df.loc[planilha_df[' ']==["Uber",f"{mês}"]]
                planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]=planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]-planilha_df.loc[planilha_df[' ']==["Saúde",f"{mês}"]]
                planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]=planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]-planilha_df.loc[planilha_df[' ']==["Streaming",f"{mês}"]]
                planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]=planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]-planilha_df.loc[planilha_df[' ']==["Igreja",f"{mês}"]]
                planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]=planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]-planilha_df.loc[planilha_df[' ']==["Lazer",f"{mês}"]]
                planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]=planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]-planilha_df.loc[planilha_df[' ']==["Roupa",f"{mês}"]]
                planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]=planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]-planilha_df.loc[planilha_df[' ']==["Comida",f"{mês}"]]

                
                saldo_final=planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]].values[0]
                bot.send_message(mensagem.chat.id,f"Então, o seu saldo final foi {saldo_final}")
                
                planilha_df.to_excel(f"{nome}.xlsx",index=False)
                

                
            else:
                nome=mensagem.from_user.first_name
                planilha_df=pd.read_excel(f"{nome}.xlsx")
                despesa_antiga=planilha_df.loc[planilha_df[' ']==f"{gasto}",f"{mês}"].values[0]
                mensagem=bot.send_message(mensagem.chat.id,f"Esse é saldo que estava antes de você querer alterar. Você tem certeza que quer trocar ?\n/Sim.\n/Não.")
                bot.register_next_step_handler(mensagem,validar_resposta)
                valor_para_trocar=bot.register_next_step_handler(mensagem,valor_correto)
                planilha_df.loc[planilha_df[' ']==[f"{gasto}",f"{mês}"]]=valor_para_trocar
                planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]=planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]+despesa_antiga
                planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]=planilha_df.loc[planilha_df[' ']==["Saldo final",f"{mês}"]]-valor_para_trocar
                
                planilha_df.to_excel(f"{nome}.xlsx",index=False)
    
    def descobrir_mês(mensagem):
        mês= mensagem.text.replace('/', '')
        texto=f'''Perfeito, então foi no mês {mês}, foi um erro em qual estrutura ?
/Saldo_depositado

/Saúde
/Comida
/Uber
/Gasolina
/Academia
/Lazer
/Roupa
/Streaming
/Igreja
/Despesa_adicional'''
        bot.send_message(mensagem.chat.id,texto)
        bot.register_next_step_handler(mensagem,descobrir_erro,mês)

    
    
    def validar_resposta(mensagem):
        resposta=mensagem.text.replace('/', '')
        if (resposta=='Sim'):
            mensagem=bot.send_message(mensagem.chat.id, "Qual valor deve ser colocado então ?")
        else:
            mensagem=bot.send_message(mensagem.chat.id, "Certo, então, para reiniciar o bot aperte no start.\n/start")
    def valor_correto(mensagem):
        return mensagem.text.replace('R$', '')
        
    texto='''Que pena que você cometeu um erro, vamos corrigir. Seu erro foi em que mês ?
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
/Dezembro'''
        
    mensagem=bot.send_message(mensagem.chat.id, texto)
    bot.register_next_step_handler(mensagem,descobrir_mês)
    
    
                
                
                
                
                
def cadastro(mensagem):
    email=mensagem.text
    bot.send_message(mensagem.chat.id, f"Seu email é {email}")
    ''' saber nome da pessoa por id'''
    finanças={
        " ":["Saldo depositado","Saldo final","Academia","Uber","Saúde","Streaming","Igreja","Roupa","Lazer","Observações","Comida","Gasolina","Despesas adicionais"],
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
