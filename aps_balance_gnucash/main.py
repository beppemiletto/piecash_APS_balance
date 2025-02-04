import datetime
import re
import piecash
from __init__ import BalanceTable, ExcelBalanceTable
import csv
from decimal import Decimal
from pathlib import Path
import pickle


def define_balance_table():
    bt = BalanceTable()
    return bt

def main():
    # choice of most updated .gnucash file from accounting directory
    # file must be saved in sqlite3 format

    gnucash_file_dir = Path('C:\\Users\\GMiletto\\Downloads')
    gnucash_file_pat = r'*.gnucash'

    try:
        latest_gnucash_file = max(gnucash_file_dir.glob(gnucash_file_pat), key=lambda f: f.stat().st_ctime)
    except FileNotFoundError:
        latest_gnucash_file = None
    except UnboundLocalError:
        latest_gnucash_file = None
    except:
        latest_gnucash_file = None




    try:
        book = piecash.open_book(latest_gnucash_file, readonly=True)
    except:
        print('File {} non aperto. Possibile errore formato sqlite3'.format(latest_gnucash_file))
        exit(2)
    finally:
        print("Aperto book uri = '{}'".format(book.uri))
    bt = define_balance_table()

    bt.GNUCASH_FILE = book.uri

    # Definizone del periodo contabile per generazione del report
    bal_period_begin_n = datetime.date(2024, 1, 1)
    bal_period_end_n = datetime.date(2024, 12, 31)
    bal_period_begin_n_1 = datetime.date(2023, 1, 1)
    bal_period_end_n_1 = datetime.date(2023, 12, 31)

    # Aggiornamento dizionario balance table bt
    bt.PERIODS['period_n']['begin'] = bal_period_begin_n
    bt.PERIODS['period_n']['end'] = bal_period_end_n
    bt.PERIODS['period_n_1']['begin'] = bal_period_begin_n_1
    bt.PERIODS['period_n_1']['end'] = bal_period_end_n_1


    # Prepara l'avanzo / disavanzo per ciascun conto del libro contabile nel periodo di bilancio
    avanzo_conti_periodo_n = {}
    avanzo_conti_periodo_n_1 = {}
    for acc in book.accounts:
        nome_conto = acc.fullname
        children_n = len(acc.children)
        acc_balance_begin_n = acc.get_balance(at_date=bal_period_begin_n)
        acc_balance_end_n = acc.get_balance(at_date=bal_period_end_n)
        acc_surplus_n = acc_balance_end_n - acc_balance_begin_n
        avanzo_conti_periodo_n[nome_conto] = [acc_surplus_n, acc_balance_begin_n, acc_balance_end_n, children_n]
        acc_balance_begin_n_1 = acc.get_balance(at_date=bal_period_begin_n_1)
        acc_balance_end_n_1 = acc.get_balance(at_date=bal_period_end_n_1)
        acc_surplus_n_1 = acc_balance_end_n_1 - acc_balance_begin_n_1
        avanzo_conti_periodo_n_1[nome_conto] = [acc_surplus_n_1, acc_balance_begin_n_1, acc_balance_end_n_1, children_n]


    del nome_conto, acc_balance_begin_n, acc_balance_end_n, acc_surplus_n, children_n
    del acc_balance_begin_n_1, acc_balance_end_n_1, acc_surplus_n_1

    # Verifica presenza conti orfani, ossia con valori in bilancio ma non aggregati a voci del
    # rendiconto per cassa: in __init__.py vengono stabilite le aggregazioni dei conti nella voci del rendiconto per
    # cassa.  La procedura seguente cerca appariglio tra tutte le voci presenti in __init__.py e quelle presenti nel book






    # Verifica uso dei conti: copertura del libro con assegnazione ai campi del bilancio
    # e calcolo dei saldi e disavanzi totali per le voci incluse
    verifica_copertura_conti = {}
    for acc in avanzo_conti_periodo_n.keys():
        account_found = 0
        places = []
        if avanzo_conti_periodo_n[acc][3] < 1:
            # trovato conto non placeholder con valori effettivi
            for sezione in bt.TABLE_BODY:
                if sezione == 'Patrimonio':
                    for voce in bt.TABLE_BODY[sezione]['DARE'].keys():
                        if type(voce) is not int:
                            continue
                        else:
                            if acc in bt.TABLE_BODY[sezione]['DARE'][voce]['accounts']:
                                account_found += 1
                                place = "Sezione {} - DARE - {}".format(sezione, voce)
                                places.append(place)
                                verifica_copertura_conti[acc] = {'found': account_found, 'places': places}
                else:
                    for voce in bt.TABLE_BODY[sezione]['USCITE'].keys():
                        if type(voce) is not int:
                            continue
                        else:
                            if acc in bt.TABLE_BODY[sezione]['USCITE'][voce]['accounts']:
                                account_found += 1
                                place = "Sezione {} - USCITE - {}".format(sezione, voce)
                                places.append(place)
                                verifica_copertura_conti[acc] = {'found': account_found, 'places': places}
                    for voce in bt.TABLE_BODY[sezione]['ENTRATE'].keys():
                        if type(voce) is not int:
                            continue
                        else:
                            if acc in bt.TABLE_BODY[sezione]['ENTRATE'][voce]['accounts']:
                                account_found += 1
                                place = "Sezione {} - ENTRATE - {}".format(sezione, voce)
                                places.append(place)
                                verifica_copertura_conti[acc] = {'found': account_found, 'places': places}


    # Calcolo delle voci in tabella secondo schema tabella bilancio APS
    GTU_Balance_n = 0
    GTU_Balance_n_1 = 0
    GTE_Balance_n = 0
    GTE_Balance_n_1 = 0
    for sezione in bt.TABLE_BODY:
        if sezione == 'Patrimonio':
            GT = Decimal(0.00)
            for voce in bt.TABLE_BODY[sezione]['DARE'].keys():
                if type(voce) is not int:
                    continue
                else:
                    totale_n = Decimal(0.00)
                    for acc in bt.TABLE_BODY[sezione]['DARE'][voce]['accounts']:
                        totale_n += avanzo_conti_periodo_n[acc][2]
                        # print(avanzo_conti_periodo_n[acc][0], totale)
                    print('Saldo finale per voce {} della sezione {} = € {}'.format(acc, sezione, totale_n))
                    print()
                    bt.TABLE_BODY[sezione]['DARE'][voce]['value_n'] = totale_n
                    GT += totale_n
            print('GTotale saldo della sezione {} = € {}'.format(sezione, GT))
            print()
            print()
            bt.TABLE_BODY[sezione]['DARE']['GT'] = {}
            bt.TABLE_BODY[sezione]['DARE']['GT']['value_n'] = GT
            bt.TABLE_BODY[sezione]['DARE']['GT']['value_n_1'] = Decimal(0.00)
            del GT, totale_n
        elif sezione in ['A', 'B', 'C', 'D', 'E', 'F']:
            GTU_n = Decimal(0.00)
            GTU_n_1 = Decimal(0.00)
            for voce in bt.TABLE_BODY[sezione]['USCITE'].keys():
                if type(voce) is not int:
                    continue
                else:
                    totale_n = Decimal(0.00)
                    totale_n_1 = Decimal(0.00)
                    for acc in bt.TABLE_BODY[sezione]['USCITE'][voce]['accounts']:
                        totale_n += avanzo_conti_periodo_n[acc][0]
                        totale_n_1 += avanzo_conti_periodo_n_1[acc][0]
                        # print(avanzo_conti_periodo_n[acc][0], totale)
                    bt.TABLE_BODY[sezione]['USCITE'][voce]['value_n'] = totale_n
                    bt.TABLE_BODY[sezione]['USCITE'][voce]['value_n_1'] = totale_n_1
                    GTU_n += totale_n
                    GTU_n_1 += totale_n_1
            bt.TABLE_BODY[sezione]['USCITE']['GT'] = {}
            bt.TABLE_BODY[sezione]['USCITE']['GT']['value_n'] = GTU_n
            bt.TABLE_BODY[sezione]['USCITE']['GT']['value_n_1'] = GTU_n_1
            # esclusione della sezione F dal totale GTU_Balance in quanto trattato successivamente a parte
            if sezione in ['A', 'B', 'C', 'D', 'E']:
                GTU_Balance_n += GTU_n
                GTU_Balance_n_1 += GTU_n_1
            del totale_n, totale_n_1

            GTE_n = Decimal(0.00)
            GTE_n_1 = Decimal(0.00)
            for voce in bt.TABLE_BODY[sezione]['ENTRATE'].keys():
                if type(voce) is not int:
                    continue
                else:
                    totale_n = Decimal(0.00)
                    totale_n_1 = Decimal(0.00)
                    for acc in bt.TABLE_BODY[sezione]['ENTRATE'][voce]['accounts']:
                        totale_n += avanzo_conti_periodo_n[acc][0]
                        totale_n_1 += avanzo_conti_periodo_n_1[acc][0]
                    bt.TABLE_BODY[sezione]['ENTRATE'][voce]['value_n'] = totale_n
                    bt.TABLE_BODY[sezione]['ENTRATE'][voce]['value_n_1'] = totale_n_1
                    GTE_n += totale_n
                    GTE_n_1 += totale_n_1
            bt.TABLE_BODY[sezione]['ENTRATE']['GT'] = {}
            bt.TABLE_BODY[sezione]['ENTRATE']['GT']['value_n'] = GTE_n
            bt.TABLE_BODY[sezione]['ENTRATE']['GT']['value_n_1'] = GTE_n_1
            # esclusione della sezione F dal totale GTE_Balance in quanto trattato successivamente a parte
            if sezione in ['A', 'B', 'C', 'D', 'E']:
                GTE_Balance_n += GTE_n
                GTE_Balance_n_1 += GTE_n_1
            del totale_n, totale_n_1
        if sezione in ['A', 'B', 'C', 'D']:
            surplus_section_n = GTE_n - GTU_n
            surplus_section_n_1 = GTE_n_1 - GTU_n_1
            bt.TABLE_BODY[sezione]['AV_DIS']['value_n'] = surplus_section_n
            bt.TABLE_BODY[sezione]['AV_DIS']['value_n_1'] = surplus_section_n_1
    surplus_Balance_n = GTE_Balance_n - GTU_Balance_n
    surplus_Balance_n_1 = GTE_Balance_n_1 - GTU_Balance_n_1
    bt.BALANCE['GTU']['value_n'] = GTU_Balance_n
    bt.BALANCE['GTE']['value_n'] = GTE_Balance_n
    bt.BALANCE['surplus_Balance']['value_n'] = surplus_Balance_n
    bt.BALANCE['GTU']['value_n_1'] = GTU_Balance_n_1
    bt.BALANCE['GTE']['value_n_1'] = GTE_Balance_n_1
    bt.BALANCE['surplus_Balance']['value_n_1'] = surplus_Balance_n_1

    del GTE_n, GTE_Balance_n, GTU_n, GTU_Balance_n, acc, GTU_n_1, GTE_n_1,  GTE_Balance_n_1, GTU_Balance_n_1

    # Salva un file csv con situazione copertura conti programma di contabilità 'verifica_conti.csv'
    with open('verifica_conti.csv', 'w', newline='') as fp:
        csv_writer = csv.writer(fp)
        for key, item in verifica_copertura_conti.items():
            row = [key, item['found']]
            for place in item['places']:
                row.append(place)
            csv_writer.writerow(row)


    # ========================= RICERCA FLUSSI CASSA PER ATTIVITA' / PASSIVITA'
    # Inizializzazione dei pattern di ricerca per modificare la contabilità gestionale in
    #  semplificazione bilancio rendiconto per cassa

    # ================================================================================= CESPITI
    # Costi per investimenti in cespiti:
    # 1) estrazione costi acquisto che entrano nel bilancio rendiconto per cassa
    # 2) flusso finanziario dei deprezzamenti - non usato nel bilancio rendiconto per cassa

    pattern_costo = re.compile(r'^Attività:Beni Cespiti:.+:[C|c]osto') # ricerca costi_cespiti_n allocati nell'anno
    pattern_depre = re.compile(r'^Uscite:[D|d]eprezzamento') # ricerca deprezzamenti costi_cespiti_n nell'anno

    #Inizializzazione dei contatori overall e del periodo contabile specifico el bilancio
    counter_ovl = 0
    counter_bal = 0

    #Inizializzazione del dictionary delle transazioni di competenza periodo di bilancio
    bal_period_transactions = {}

    # Ricerca delle transazioni del periodo di bilancio relative a costi da mettere in bilancio per cassa
    # ma che non sarebbero di flusso di cassa (costi acquisto costi_cespiti_n) che vanno normalmente negli investimenti

    # Costi sostenuti per acquisto cespiti nell'esercizio corrente n
    costi_cespiti_n = Decimal(0.00)
    depre_cespiti_n = Decimal(0.00)
    for transaction in book.transactions:
        counter_ovl += 1
        post_date = transaction.post_date
        # verifica che la data  post dello split sia inclusa nel periodo del bilancio
        if post_date >= bal_period_begin_n and post_date <= bal_period_end_n:
            counter_bal += 1
            bal_period_transactions[str(counter_bal)] = {
                'date': transaction.post_date,
                'descr': transaction.description,
                'splits': transaction.splits
            }

            for split in transaction.splits:
                acc_book_fullname = split.account.fullname
                # print(fullname)
                reg_res_costo = re.match(pattern_costo, acc_book_fullname)
                reg_res_depre = re.match(pattern_depre, acc_book_fullname)
                # print(reg_res)
                if reg_res_costo is not None:
                    costi_cespiti_n += split.value

                # print(reg_res)
                elif reg_res_depre is not None:
                    depre_cespiti_n += split.value


    del bal_period_transactions

    #Inizializzazione dei contatori overall e del periodo contabile specifico el bilancio
    counter_ovl = 0
    counter_bal = 0

    #Inizializzazione del dictionary delle transazioni di competenza periodo di bilancio
    bal_period_transactions = {}


    # Costi sostenuti per acquisto cespiti nell'esercizio precedente n-1
    costi_cespiti_n_1 = Decimal(0.00)
    depre_cespiti_n_1 = Decimal(0.00)
    for transaction in book.transactions:
        counter_ovl += 1
        post_date = transaction.post_date
        # verifica che la data  post dello split sia inclusa nel periodo del bilancio
        if post_date >= bal_period_begin_n_1 and post_date <= bal_period_end_n_1:
            counter_bal += 1
            bal_period_transactions[str(counter_bal)] = {
                'date': transaction.post_date,
                'descr': transaction.description,
                'splits': transaction.splits
            }

            for split in transaction.splits:
                acc_book_fullname = split.account.fullname
                # print(fullname)
                reg_res_costo = re.match(pattern_costo, acc_book_fullname)
                reg_res_depre = re.match(pattern_depre, acc_book_fullname)
                # print(reg_res)
                if reg_res_costo is not None:
                    costi_cespiti_n_1 += split.value
                # print(reg_res)
                elif reg_res_depre is not None:
                    depre_cespiti_n_1 += split.value


    bt.ASSETT['COSTO']['value_n'] = costi_cespiti_n
    bt.ASSETT['TABLE']['USCITE'][1]['value_n'] = costi_cespiti_n
    bt.ASSETT['TABLE']['USCITE'][1]['value_n_1'] = costi_cespiti_n_1
    bt.ASSETT['COSTO']['value_n_1'] = costi_cespiti_n_1
    bt.ASSETT['DEPREZZAMENTO']['value_n'] = depre_cespiti_n
    bt.ASSETT['DEPREZZAMENTO']['value_n_1'] = depre_cespiti_n_1
    # duplicazione del valore nella sezione F utilizzata per stampare il foglio Excel
    bt.TABLE_BODY['F']['USCITE'][1]['value_n'] = costi_cespiti_n
    bt.TABLE_BODY['F']['USCITE'][1]['value_n_1'] = costi_cespiti_n_1

    del costi_cespiti_n, depre_cespiti_n, reg_res_costo, reg_res_depre
    del counter_bal, counter_ovl, costi_cespiti_n_1, depre_cespiti_n_1

    # Prestiti all'Associazione sotto forma di anticipi da parte dei soci e Restituzione di prestiti

    # ====================================================================== PRESTITI E RESTITUZIONI
    # RICERCA DI prestiti a LTC come anticipo spese o prestito non fruttifero:
    # RICERCA DELLE TRANSAZIONI IN USCITA DAI CONTI PASSIVITA' COME RESITUZIONI

    pattern_prestito = re.compile(r'^Passività:Anticipi spese da soci:.+')  # ricerca conti singoli soci allocati nell'anno

    # Inizializzazione dei contatori overall e del periodo contabile specifico el bilancio
    counter_ovl = 0
    counter_bal = 0

    # Inizializzazione del dictionary delle transazioni di competenza periodo di bilancio
    bal_period_transactions = {}

    # Ricerca delle transazioni del periodo di bilancio relative a entrate  da mettere in bilancio per cassa
    # ma che non sarebbero di flusso di cassa (passitita) che vanno normalmente nello stato patrimoniale

    # Entrate prestito soci nell'esercizio corrente n
    prestiti_n = Decimal(0.00)
    restituzione_n = Decimal(0.00)
    for transaction in book.transactions:
        counter_ovl += 1
        post_date = transaction.post_date
        # verifica che la data  post dello split sia inclusa nel periodo del bilancio
        if post_date >= bal_period_begin_n and post_date <= bal_period_end_n:
            counter_bal += 1
            bal_period_transactions[str(counter_bal)] = {
                'date': transaction.post_date,
                'descr': transaction.description,
                'splits': transaction.splits
            }

            for split in transaction.splits:
                acc_book_fullname = split.account.fullname
                # print(fullname)
                reg_res_prestito = re.match(pattern_prestito, acc_book_fullname)
                # print(reg_res)
                if reg_res_prestito is not None:
                    if split.value < 0:
                        prestiti_n += abs(split.value)
                    elif split.value > 0:
                        restituzione_n += abs(split.value)



    del bal_period_transactions

    # Inizializzazione dei contatori overall e del periodo contabile specifico el bilancio
    counter_ovl = 0
    counter_bal = 0

    # Inizializzazione del dictionary delle transazioni di competenza periodo di bilancio
    bal_period_transactions = {}

    # Prestiti e Restituzioni prestiti nell'esercizio precedente n-1
    prestiti_n_1 = Decimal(0.00)
    restituzione_n_1 = Decimal(0.00)
    for transaction in book.transactions:
        counter_ovl += 1
        post_date = transaction.post_date
        # verifica che la data  post dello split sia inclusa nel periodo del bilancio
        if post_date >= bal_period_begin_n_1 and post_date <= bal_period_end_n_1:
            counter_bal += 1
            bal_period_transactions[str(counter_bal)] = {
                'date': transaction.post_date,
                'descr': transaction.description,
                'splits': transaction.splits
            }

            for split in transaction.splits:
                acc_book_fullname = split.account.fullname
                # print(fullname)
                reg_res_prestito = re.match(pattern_prestito, acc_book_fullname)
                if reg_res_prestito is not None:
                    if split.value < 0:
                        prestiti_n_1 += abs(split.value)
                    elif split.value > 0:
                        restituzione_n_1 += abs(split.value)




    bt.ASSETT['PRESTITI']['value_n'] = avanzo_conti_periodo_n['Passività:Anticipi spese da soci'][0]
    bt.ASSETT['PRESTITI']['value_n_1'] = avanzo_conti_periodo_n_1['Passività:Anticipi spese da soci'][0]
    bt.ASSETT['TABLE']['ENTRATE'][4]['value_n'] = avanzo_conti_periodo_n['Passività:Anticipi spese da soci'][0]
    bt.ASSETT['TABLE']['ENTRATE'][4]['value_n_1'] = avanzo_conti_periodo_n_1['Passività:Anticipi spese da soci'][0]

    # duplicazione del valore nella sezione F utilizzata per stampare il foglio Excel
    bt.TABLE_BODY['F']['ENTRATE'][4]['value_n'] = prestiti_n
    bt.TABLE_BODY['F']['ENTRATE'][4]['value_n_1'] = prestiti_n_1
    bt.TABLE_BODY['F']['USCITE'][4]['value_n'] = restituzione_n
    bt.TABLE_BODY['F']['USCITE'][4]['value_n_1'] = restituzione_n_1

    del prestiti_n, prestiti_n_1, reg_res_prestito, bal_period_transactions
    del counter_bal, counter_ovl, restituzione_n, restituzione_n_1

    # estrazione dei saldi Cassa e Banca
    cash = {}
    cash['accounts']= bt.TABLE_BODY['Patrimonio']['DARE'][1]['accounts']
    cash_balance_n = 0
    cash_balance_n_1 = 0
    for acc in cash['accounts']:
        cash_balance_n += avanzo_conti_periodo_n[acc][2]
        cash_balance_n_1 += avanzo_conti_periodo_n_1[acc][2]
    bt.ASSETT['CASSA']['value_n'] = cash_balance_n
    bt.ASSETT['CASSA']['value_n_1'] = cash_balance_n_1

    bank = {}
    bank['accounts'] = bt.TABLE_BODY['Patrimonio']['DARE'][2]['accounts']
    bank_balance_n = 0
    bank_balance_n_1 = 0
    for acc in bank['accounts']:
        bank_balance_n += avanzo_conti_periodo_n[acc][2]
        bank_balance_n_1 += avanzo_conti_periodo_n_1[acc][2]
    bt.ASSETT['BANCA']['value_n'] = bank_balance_n
    bt.ASSETT['BANCA']['value_n_1'] = bank_balance_n_1

    del bank, bank_balance_n, bank_balance_n_1, cash, cash_balance_n, cash_balance_n_1

    # Calcolo delle imposte specifiche sul reddito IRES e IRAP
    # per voce in quadro Imposte alla rine del quadro E (riga file excel 55)

    IRESeIRAP_n = avanzo_conti_periodo_n['Uscite:Tasse:IRES e IRAP'][0]
    IRESeIRAP_n_1 = avanzo_conti_periodo_n_1['Uscite:Tasse:IRES e IRAP'][0]

    # La classe BalanceTable() contiene le righe di formattazione
    # del foglio Excel sul quale viene scritta la tabella. Il foglio Excel viene scritto come file
    # Excel con formati e metodi definiti nella classe ExcelBalanceTable() che si basa sul
    # modulo Python openpyxl.styles import Font, Color

    # =========================================================================================================
    # Generazione del report Excel macro blocco
    # =========================================================================================================
    if True:
        # Scrittura del foglio Excel del Bilancio con classe
        etb = ExcelBalanceTable("Bilancio_LaboratorioTeatraleDiCambianoAPS_{}.xls".format(bal_period_end_n.year))

        # Page Header
        data_line = [1, 1, [bt.HEADER.format(bal_period_end_n.year)],6]
        print(data_line)
        cella = etb.writeline(dataline=data_line, bold=True, wrap=True, fontsize=14, halign='center')

        # Page Title
        data_line = [1, 2, [bt.TITLE],6]
        print(data_line)
        cella = etb.writeline(dataline=data_line, bold=True, wrap=True, fontsize=16, halign='center')

        # Page SubTitle
        data_line = [1, 3, [bt.SUBTITLE],6]
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=26, italic=True, wrap=True, fontsize=8, halign='left')

        # Page Blank small line
        data_line = [1, 4, ["  "],6]
        cella = etb.writeline(dataline=data_line, row_height=2, italic=True, wrap=True, fontsize=6, halign='left')

        # Table Header TOP
        data_line = [1,5]
        data = [bt.TABLE_HEADER[0]]
        data.append(bt.TABLE_HEADER[1].format(n=bal_period_end_n.year))
        data.append(bt.TABLE_HEADER[2].format(n_1=bal_period_end_n_1.year))
        data.append(bt.TABLE_HEADER[3])
        data.append(bt.TABLE_HEADER[4].format(n=bal_period_end_n.year))
        data.append(bt.TABLE_HEADER[5].format(n_1=bal_period_end_n_1.year))
        data_line.append(data)
        data_line.append(None)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=15, bold=True, wrap=False, fontsize=11, halign='center', border=True)

        # Sezione A
        # Sezione A - Header

        data_line = [1,6]
        data = [bt.TABLE_BODY['A']['USCITE']['title']]
        data.append(bt.TABLE_HEADER[1].format(n=bal_period_end_n.year))
        data.append(bt.TABLE_HEADER[2].format(n_1=bal_period_end_n_1.year))
        data.append(bt.TABLE_BODY['A']['ENTRATE']['title'])
        data.append(bt.TABLE_HEADER[4].format(n=bal_period_end_n.year))
        data.append(bt.TABLE_HEADER[5].format(n_1=bal_period_end_n_1.year))
        data_line.append(data)
        data_line.append(None)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=20, bold=True, wrap=True, fontsize=9, halign='left', border=True)

        # Sezione A - Uscite

        if True:
            # Sezione A - Uscite - 1
            data_line = [1,8]
            data = [bt.TABLE_BODY['A']['USCITE'][1]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left', border=True)
            data_line = [2,8]
            data = [bt.TABLE_BODY['A']['USCITE'][1]['value_n']]
            data.append(bt.TABLE_BODY['A']['USCITE'][1]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione A - Uscite - 2
            data_line = [1,9]
            data = [bt.TABLE_BODY['A']['USCITE'][2]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left', border=True)
            data_line = [2,9]
            data = [bt.TABLE_BODY['A']['USCITE'][2]['value_n']]
            data.append(bt.TABLE_BODY['A']['USCITE'][2]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right', border=True)

            # Sezione A - Uscite - 3
            data_line = [1,11]
            data = [bt.TABLE_BODY['A']['USCITE'][3]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left', border=True)
            data_line = [2,11]
            data = [bt.TABLE_BODY['A']['USCITE'][3]['value_n']]
            data.append(bt.TABLE_BODY['A']['USCITE'][3]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right', border=True)

            # Sezione A - Uscite - 4
            data_line = [1,13]
            data = [bt.TABLE_BODY['A']['USCITE'][4]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left', border=True)
            data_line = [2,13]
            data = [bt.TABLE_BODY['A']['USCITE'][4]['value_n']]
            data.append(bt.TABLE_BODY['A']['USCITE'][4]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right', border=True)

            # Sezione A - Uscite - 5
            data_line = [1, 15]
            data = [bt.TABLE_BODY['A']['USCITE'][5]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left', border=True)
            data_line = [2, 15]
            data = [bt.TABLE_BODY['A']['USCITE'][5]['value_n']]
            data.append(bt.TABLE_BODY['A']['USCITE'][5]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right', border=True)

            # Sezione A - Uscite - Totalizzatore
            data_line = [1, 17]
            data = ['Totale']
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='right', border=True)
            data_line = [2, 17]
            data = [bt.TABLE_BODY['A']['USCITE']['GT']['value_n']]
            data.append(bt.TABLE_BODY['A']['USCITE']['GT']['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right', border=True)

        # Sezione A - Entrate

        if True:
            # Sezione A - Entrate - 1
            data_line = [4, 7]
            data = [bt.TABLE_BODY['A']['ENTRATE'][1]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 7]
            data = [bt.TABLE_BODY['A']['ENTRATE'][1]['value_n']]
            data.append(bt.TABLE_BODY['A']['ENTRATE'][1]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione A - Entrate - 2
            data_line = [4, 8]
            data = [bt.TABLE_BODY['A']['ENTRATE'][2]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 8]
            data = [bt.TABLE_BODY['A']['ENTRATE'][2]['value_n']]
            data.append(bt.TABLE_BODY['A']['ENTRATE'][2]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione A - Entrate - 3
            data_line = [4, 9]
            data = [bt.TABLE_BODY['A']['ENTRATE'][3]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 9]
            data = [bt.TABLE_BODY['A']['ENTRATE'][3]['value_n']]
            data.append(bt.TABLE_BODY['A']['ENTRATE'][3]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione A - Entrate - 4
            data_line = [4, 10]
            data = [bt.TABLE_BODY['A']['ENTRATE'][4]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 10]
            data = [bt.TABLE_BODY['A']['ENTRATE'][4]['value_n']]
            data.append(bt.TABLE_BODY['A']['ENTRATE'][4]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione A - Entrate - 5
            data_line = [4, 11]
            data = [bt.TABLE_BODY['A']['ENTRATE'][5]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 11]
            data = [bt.TABLE_BODY['A']['ENTRATE'][5]['value_n']]
            data.append(bt.TABLE_BODY['A']['ENTRATE'][5]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione A - Entrate - 6
            data_line = [4, 12]
            data = [bt.TABLE_BODY['A']['ENTRATE'][6]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 12]
            data = [bt.TABLE_BODY['A']['ENTRATE'][6]['value_n']]
            data.append(bt.TABLE_BODY['A']['ENTRATE'][6]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione A - Entrate - 7
            data_line = [4, 13]
            data = [bt.TABLE_BODY['A']['ENTRATE'][7]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 13]
            data = [bt.TABLE_BODY['A']['ENTRATE'][7]['value_n']]
            data.append(bt.TABLE_BODY['A']['ENTRATE'][7]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione A - Entrate - 8
            data_line = [4, 14]
            data = [bt.TABLE_BODY['A']['ENTRATE'][8]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 14]
            data = [bt.TABLE_BODY['A']['ENTRATE'][8]['value_n']]
            data.append(bt.TABLE_BODY['A']['ENTRATE'][8]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione A - Entrate - 9
            data_line = [4, 15]
            data = [bt.TABLE_BODY['A']['ENTRATE'][9]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 15]
            data = [bt.TABLE_BODY['A']['ENTRATE'][9]['value_n']]
            data.append(bt.TABLE_BODY['A']['ENTRATE'][9]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)


            # Sezione A - Entrate - 10
            data_line = [4, 16]
            data = [bt.TABLE_BODY['A']['ENTRATE'][10]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 16]
            data = [bt.TABLE_BODY['A']['ENTRATE'][10]['value_n']]
            data.append(bt.TABLE_BODY['A']['ENTRATE'][10]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione A - Entrate - Totalizzatore
            data_line = [4, 17]
            data = ['Totale']
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [5, 17]
            data = [bt.TABLE_BODY['A']['ENTRATE']['GT']['value_n']]
            data.append(bt.TABLE_BODY['A']['ENTRATE']['GT']['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione A - Totalizzatore avanzo
            data_line = [4, 18]
            data = [bt.TABLE_BODY['A']['AV_DIS']['title']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [5, 18]
            data = [bt.TABLE_BODY['A']['AV_DIS']['value_n']]
            data.append(bt.TABLE_BODY['A']['AV_DIS']['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

        # Page Blank small line
        data_line = [1, 19, ["  "],6]
        cella = etb.writeline(dataline=data_line, row_height=2, italic=True, wrap=True, fontsize=6, halign='left')

        # Sezione B
        # Sezione B - Header

        data_line = [1, 20]
        data = [bt.TABLE_BODY['B']['USCITE']['title']]
        data.append(bt.TABLE_HEADER[1].format(n=bal_period_end_n.year))
        data.append(bt.TABLE_HEADER[2].format(n_1=bal_period_end_n_1.year))
        data.append(bt.TABLE_BODY['B']['ENTRATE']['title'])
        data.append(bt.TABLE_HEADER[4].format(n=bal_period_end_n.year))
        data.append(bt.TABLE_HEADER[5].format(n_1=bal_period_end_n_1.year))
        data_line.append(data)
        data_line.append(None)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=20, bold=True, wrap=True, fontsize=9, halign='left',
                              border=True)

        # Sezione B - Uscite

        if True:
            # Sezione B - Uscite - 1
            data_line = [1, 21]
            data = [bt.TABLE_BODY['B']['USCITE'][1]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 21]
            data = [bt.TABLE_BODY['B']['USCITE'][1]['value_n']]
            data.append(bt.TABLE_BODY['B']['USCITE'][1]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione B - Uscite - 2
            data_line = [1, 22]
            data = [bt.TABLE_BODY['B']['USCITE'][2]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 22]
            data = [bt.TABLE_BODY['B']['USCITE'][2]['value_n']]
            data.append(bt.TABLE_BODY['B']['USCITE'][2]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione B - Uscite - 3
            data_line = [1, 23]
            data = [bt.TABLE_BODY['B']['USCITE'][3]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 23]
            data = [bt.TABLE_BODY['B']['USCITE'][3]['value_n']]
            data.append(bt.TABLE_BODY['B']['USCITE'][3]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione B - Uscite - 4
            data_line = [1, 24]
            data = [bt.TABLE_BODY['B']['USCITE'][4]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 24]
            data = [bt.TABLE_BODY['B']['USCITE'][4]['value_n']]
            data.append(bt.TABLE_BODY['B']['USCITE'][4]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione B - Uscite - 5
            data_line = [1, 25]
            data = [bt.TABLE_BODY['B']['USCITE'][5]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 25]
            data = [bt.TABLE_BODY['B']['USCITE'][5]['value_n']]
            data.append(bt.TABLE_BODY['B']['USCITE'][5]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione B - Uscite - Totalizzatore
            data_line = [1, 27]
            data = ['Totale']
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [2, 27]
            data = [bt.TABLE_BODY['B']['USCITE']['GT']['value_n']]
            data.append(bt.TABLE_BODY['B']['USCITE']['GT']['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

        # Sezione B - Entrate

        if True:
            # Sezione B - Entrate - 1
            data_line = [4, 21]
            data = [bt.TABLE_BODY['B']['ENTRATE'][1]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 21]
            data = [bt.TABLE_BODY['B']['ENTRATE'][1]['value_n']]
            data.append(bt.TABLE_BODY['B']['ENTRATE'][1]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione B - Entrate - 2
            data_line = [4, 22]
            data = [bt.TABLE_BODY['B']['ENTRATE'][2]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 22]
            data = [bt.TABLE_BODY['B']['ENTRATE'][2]['value_n']]
            data.append(bt.TABLE_BODY['B']['ENTRATE'][2]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione B - Entrate - 3
            data_line = [4, 23]
            data = [bt.TABLE_BODY['B']['ENTRATE'][3]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 23]
            data = [bt.TABLE_BODY['B']['ENTRATE'][3]['value_n']]
            data.append(bt.TABLE_BODY['B']['ENTRATE'][3]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione B - Entrate - 4
            data_line = [4, 24]
            data = [bt.TABLE_BODY['B']['ENTRATE'][4]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 24]
            data = [bt.TABLE_BODY['B']['ENTRATE'][4]['value_n']]
            data.append(bt.TABLE_BODY['B']['ENTRATE'][4]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione B - Entrate - 5
            data_line = [4, 25]
            data = [bt.TABLE_BODY['B']['ENTRATE'][5]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 25]
            data = [bt.TABLE_BODY['B']['ENTRATE'][5]['value_n']]
            data.append(bt.TABLE_BODY['B']['ENTRATE'][5]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione B - Entrate - 6
            data_line = [4, 26]
            data = [bt.TABLE_BODY['B']['ENTRATE'][6]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 26]
            data = [bt.TABLE_BODY['B']['ENTRATE'][6]['value_n']]
            data.append(bt.TABLE_BODY['B']['ENTRATE'][6]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)


            # Sezione B - Entrate - Totalizzatore
            data_line = [4, 27]
            data = ['Totale']
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [5, 27]
            data = [bt.TABLE_BODY['B']['ENTRATE']['GT']['value_n']]
            data.append(bt.TABLE_BODY['B']['ENTRATE']['GT']['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione B - Totalizzatore avanzo
            data_line = [4, 28]
            data = [bt.TABLE_BODY['B']['AV_DIS']['title']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [5, 28]
            data = [bt.TABLE_BODY['B']['AV_DIS']['value_n']]
            data.append(bt.TABLE_BODY['B']['AV_DIS']['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)


        # Page Blank small line
        data_line = [1, 29, ["  "],6]
        cella = etb.writeline(dataline=data_line, row_height=2, italic=True, wrap=True, fontsize=6, halign='left')

        # Sezione C
        # Sezione C - Header

        data_line = [1, 30]
        data = [bt.TABLE_BODY['C']['USCITE']['title']]
        data.append(bt.TABLE_HEADER[1].format(n=bal_period_end_n.year))
        data.append(bt.TABLE_HEADER[2].format(n_1=bal_period_end_n_1.year))
        data.append(bt.TABLE_BODY['C']['ENTRATE']['title'])
        data.append(bt.TABLE_HEADER[4].format(n=bal_period_end_n.year))
        data.append(bt.TABLE_HEADER[5].format(n_1=bal_period_end_n_1.year))
        data_line.append(data)
        data_line.append(None)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=20, bold=True, wrap=True, fontsize=9, halign='left',
                              border=True)

        # Sezione C - Uscite

        if True:
            # Sezione C - Uscite - 1
            data_line = [1, 31]
            data = [bt.TABLE_BODY['C']['USCITE'][1]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 31]
            data = [bt.TABLE_BODY['C']['USCITE'][1]['value_n']]
            data.append(bt.TABLE_BODY['C']['USCITE'][1]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione C - Uscite - 2
            data_line = [1, 32]
            data = [bt.TABLE_BODY['C']['USCITE'][2]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 32]
            data = [bt.TABLE_BODY['C']['USCITE'][2]['value_n']]
            data.append(bt.TABLE_BODY['C']['USCITE'][2]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione C - Uscite - 3
            data_line = [1, 33]
            data = [bt.TABLE_BODY['C']['USCITE'][3]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 33]
            data = [bt.TABLE_BODY['C']['USCITE'][3]['value_n']]
            data.append(bt.TABLE_BODY['C']['USCITE'][3]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione C - Uscite - Totalizzatore
            data_line = [1, 34]
            data = ['Totale']
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [2, 34]
            data = [bt.TABLE_BODY['C']['USCITE']['GT']['value_n']]
            data.append(bt.TABLE_BODY['C']['USCITE']['GT']['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

        # Sezione C - Entrate

        if True:
            # Sezione C - Entrate - 1
            data_line = [4, 31]
            data = [bt.TABLE_BODY['C']['ENTRATE'][1]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 31]
            data = [bt.TABLE_BODY['C']['ENTRATE'][1]['value_n']]
            data.append(bt.TABLE_BODY['C']['ENTRATE'][1]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione C - Entrate - 2
            data_line = [4, 32]
            data = [bt.TABLE_BODY['C']['ENTRATE'][2]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 32]
            data = [bt.TABLE_BODY['C']['ENTRATE'][2]['value_n']]
            data.append(bt.TABLE_BODY['C']['ENTRATE'][2]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione C - Entrate - 3
            data_line = [4, 33]
            data = [bt.TABLE_BODY['C']['ENTRATE'][3]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 33]
            data = [bt.TABLE_BODY['C']['ENTRATE'][3]['value_n']]
            data.append(bt.TABLE_BODY['C']['ENTRATE'][3]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione C - Entrate - Totalizzatore
            data_line = [4, 34]
            data = ['Totale']
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [5, 34]
            data = [bt.TABLE_BODY['C']['ENTRATE']['GT']['value_n']]
            data.append(bt.TABLE_BODY['C']['ENTRATE']['GT']['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione C - Totalizzatore avanzo
            data_line = [4, 35]
            data = [bt.TABLE_BODY['C']['AV_DIS']['title']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [5, 35]
            data = [bt.TABLE_BODY['C']['AV_DIS']['value_n']]
            data.append(bt.TABLE_BODY['C']['AV_DIS']['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)


        # Page Blank small line
        data_line = [1, 36, ["  "],6]
        cella = etb.writeline(dataline=data_line, row_height=2, italic=True, wrap=True, fontsize=6, halign='left')

        # Sezione D
        # Sezione D - Header

        data_line = [1, 37]
        data = [bt.TABLE_BODY['D']['USCITE']['title']]
        data.append(bt.TABLE_HEADER[1].format(n=bal_period_end_n.year))
        data.append(bt.TABLE_HEADER[2].format(n_1=bal_period_end_n_1.year))
        data.append(bt.TABLE_BODY['D']['ENTRATE']['title'])
        data.append(bt.TABLE_HEADER[4].format(n=bal_period_end_n.year))
        data.append(bt.TABLE_HEADER[5].format(n_1=bal_period_end_n_1.year))
        data_line.append(data)
        data_line.append(None)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=20, bold=True, wrap=True, fontsize=9, halign='left',
                              border=True)

        # Sezione D - Uscite

        if True:
            # Sezione D - Uscite - 1
            data_line = [1, 38]
            data = [bt.TABLE_BODY['D']['USCITE'][1]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 38]
            data = [bt.TABLE_BODY['D']['USCITE'][1]['value_n']]
            data.append(bt.TABLE_BODY['D']['USCITE'][1]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione D - Uscite - 2
            data_line = [1, 39]
            data = [bt.TABLE_BODY['D']['USCITE'][2]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 39]
            data = [bt.TABLE_BODY['D']['USCITE'][2]['value_n']]
            data.append(bt.TABLE_BODY['D']['USCITE'][2]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione D - Uscite - 3
            data_line = [1, 40]
            data = [bt.TABLE_BODY['D']['USCITE'][3]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 40]
            data = [bt.TABLE_BODY['D']['USCITE'][3]['value_n']]
            data.append(bt.TABLE_BODY['D']['USCITE'][3]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione D - Uscite - 4
            data_line = [1, 41]
            data = [bt.TABLE_BODY['D']['USCITE'][4]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 41]
            data = [bt.TABLE_BODY['D']['USCITE'][4]['value_n']]
            data.append(bt.TABLE_BODY['D']['USCITE'][4]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione D - Uscite - 5
            data_line = [1, 42]
            data = [bt.TABLE_BODY['D']['USCITE'][5]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 42]
            data = [bt.TABLE_BODY['D']['USCITE'][5]['value_n']]
            data.append(bt.TABLE_BODY['D']['USCITE'][5]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione D - Uscite - Totalizzatore
            data_line = [1, 43]
            data = ['Totale']
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [2, 43]
            data = [bt.TABLE_BODY['D']['USCITE']['GT']['value_n']]
            data.append(bt.TABLE_BODY['D']['USCITE']['GT']['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

        # Sezione D - Entrate

        if True:
            # Sezione D - Entrate - 1
            data_line = [4, 38]
            data = [bt.TABLE_BODY['D']['ENTRATE'][1]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 38]
            data = [bt.TABLE_BODY['D']['ENTRATE'][1]['value_n']]
            data.append(bt.TABLE_BODY['D']['ENTRATE'][1]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione D - Entrate - 2
            data_line = [4, 39]
            data = [bt.TABLE_BODY['D']['ENTRATE'][2]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 39]
            data = [bt.TABLE_BODY['D']['ENTRATE'][2]['value_n']]
            data.append(bt.TABLE_BODY['D']['ENTRATE'][2]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione D - Entrate - 3
            data_line = [4, 40]
            data = [bt.TABLE_BODY['D']['ENTRATE'][3]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 40]
            data = [bt.TABLE_BODY['D']['ENTRATE'][3]['value_n']]
            data.append(bt.TABLE_BODY['D']['ENTRATE'][3]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione D - Entrate - 4
            data_line = [4, 41]
            data = [bt.TABLE_BODY['D']['ENTRATE'][4]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 41]
            data = [bt.TABLE_BODY['D']['ENTRATE'][4]['value_n']]
            data.append(bt.TABLE_BODY['D']['ENTRATE'][4]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione D - Entrate - 5
            data_line = [4, 42]
            data = [bt.TABLE_BODY['D']['ENTRATE'][5]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 42]
            data = [bt.TABLE_BODY['D']['ENTRATE'][5]['value_n']]
            data.append(bt.TABLE_BODY['D']['ENTRATE'][5]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione D - Entrate - Totalizzatore
            data_line = [4, 43]
            data = ['Totale']
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [5, 43]
            data = [bt.TABLE_BODY['D']['ENTRATE']['GT']['value_n']]
            data.append(bt.TABLE_BODY['D']['ENTRATE']['GT']['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione D - Totalizzatore avanzo
            data_line = [4, 44]
            data = [bt.TABLE_BODY['D']['AV_DIS']['title']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [5, 44]
            data = [bt.TABLE_BODY['D']['AV_DIS']['value_n']]
            data.append(bt.TABLE_BODY['D']['AV_DIS']['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

        # Page Blank small line
        data_line = [1, 45, ["  "],6]
        cella = etb.writeline(dataline=data_line, row_height=2, italic=True, wrap=True, fontsize=6, halign='left')

        # Sezione E - Header
        # Sezione E

        data_line = [1, 46]
        data = [bt.TABLE_BODY['E']['USCITE']['title']]
        data.append(bt.TABLE_HEADER[1].format(n=bal_period_end_n.year))
        data.append(bt.TABLE_HEADER[2].format(n_1=bal_period_end_n_1.year))
        data.append(bt.TABLE_BODY['E']['ENTRATE']['title'])
        data.append(bt.TABLE_HEADER[4].format(n=bal_period_end_n.year))
        data.append(bt.TABLE_HEADER[5].format(n_1=bal_period_end_n_1.year))
        data_line.append(data)
        data_line.append(None)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=20, bold=True, wrap=True, fontsize=9, halign='left',
                              border=True)

        # Sezione E - Uscite

        if True:
            # Sezione E - Uscite - 1
            data_line = [1, 47]
            data = [bt.TABLE_BODY['E']['USCITE'][1]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 47]
            data = [bt.TABLE_BODY['E']['USCITE'][1]['value_n']]
            data.append(bt.TABLE_BODY['E']['USCITE'][1]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione E - Uscite - 2
            data_line = [1, 48]
            data = [bt.TABLE_BODY['E']['USCITE'][2]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 48]
            data = [bt.TABLE_BODY['E']['USCITE'][2]['value_n']]
            data.append(bt.TABLE_BODY['E']['USCITE'][2]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione E - Uscite - 3
            data_line = [1, 49]
            data = [bt.TABLE_BODY['E']['USCITE'][3]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 49]
            data = [bt.TABLE_BODY['E']['USCITE'][3]['value_n']]
            data.append(bt.TABLE_BODY['E']['USCITE'][3]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione E - Uscite - 4
            data_line = [1, 50]
            data = [bt.TABLE_BODY['E']['USCITE'][4]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 50]
            data = [bt.TABLE_BODY['E']['USCITE'][4]['value_n']]
            data.append(bt.TABLE_BODY['E']['USCITE'][4]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione E - Uscite - 5
            data_line = [1, 51]
            data = [bt.TABLE_BODY['E']['USCITE'][5]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 51]
            data = [bt.TABLE_BODY['E']['USCITE'][5]['value_n']]
            data.append(bt.TABLE_BODY['E']['USCITE'][5]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione E - Uscite - Totalizzatore
            data_line = [1, 52]
            data = ['Totale']
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [2, 52]
            data = [bt.TABLE_BODY['E']['USCITE']['GT']['value_n']]
            data.append(bt.TABLE_BODY['E']['USCITE']['GT']['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

        # Sezione E - Entrate

        if True:
            # Sezione E - Entrate - 1
            data_line = [4, 47]
            data = [bt.TABLE_BODY['E']['ENTRATE'][1]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 47]
            data = [bt.TABLE_BODY['E']['ENTRATE'][1]['value_n']]
            data.append(bt.TABLE_BODY['E']['ENTRATE'][1]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione E - Entrate - 2
            data_line = [4, 48]
            data = [bt.TABLE_BODY['E']['ENTRATE'][2]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 48]
            data = [bt.TABLE_BODY['E']['ENTRATE'][2]['value_n']]
            data.append(bt.TABLE_BODY['E']['ENTRATE'][2]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione E - Entrate - Totalizzatore
            data_line = [4, 52]
            data = ['Totale']
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [5, 52]
            data = [bt.TABLE_BODY['E']['ENTRATE']['GT']['value_n']]
            data.append(bt.TABLE_BODY['E']['ENTRATE']['GT']['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

        # Totali USCITE ed ENTRATE e avanzo/disavanzo di esercizio
        if True:
            # Gestione - Uscite - Totalizzatore
            data_line = [1, 53]
            data = ['Totale uscite della gestione']
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=True, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [2, 53]
            data = [bt.BALANCE['GTU']['value_n']]
            data.append(bt.BALANCE['GTU']['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=True, wrap=True, fontsize=9, halign='right',
                                  border=True)
            
            # Gestione - Entrate - Totalizzatore
            data_line = [4, 53]
            data = ['Totale entrate della gestione']
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=True, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [5, 53]
            data = [bt.BALANCE['GTE']['value_n']]
            data.append(bt.BALANCE['GTE']['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=True, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Gestione - Avanzo / disavanzo - Totalizzatore
            data_line = [4, 54]
            data = ['Avanzo/disavanzo d’esercizio prima delle imposte']
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=True, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [5, 54]
            data = [bt.BALANCE['surplus_Balance']['value_n']]
            data.append(bt.BALANCE['surplus_Balance']['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=True, wrap=True, fontsize=9, halign='right',
                                  border=True)

        # Gestione Imposte
        if True:
            # Gestione - Imposte
            data_line = [4, 55]
            data = ['Imposte']
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=True, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 55]
            data = [IRESeIRAP_n]
            data.append(IRESeIRAP_n_1)
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=True, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Gestione Avanzo Disavanzo prima di investimenti
            if True:
                # Gestione - Imposte
                data_line = [4, 56]
                data = ['Avanzo/disavanzo d’esercizio prima di investimenti e disinvestimenti patrimoniali, e finanziamenti']
                data_line.append(data)
                data_line.append(None)
                print(data_line)
                cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9,
                                      halign='left',
                                      border=True)
                data_line = [5, 56]
                data = ['=E54 - E55']
                data.append('=F54 - F55')
                data_line.append(data)
                data_line.append(None)
                print(data_line)
                cella = etb.writeline(dataline=data_line, row_height=21, bold=True, wrap=True, fontsize=9,
                                      halign='right',
                                      border=True)

        # Page Blank small line
        data_line = [1, 57, ["  "],6]
        cella = etb.writeline(dataline=data_line, row_height=2, italic=True, wrap=True, fontsize=6, halign='left')

        # Sezione Cespiti Immobilizzazioni
        # Sezione Cespiti Immobilizzazioni  - Header - detta anche sezione F per flussi di cassa

        data_line = [1, 58]
        data = [bt.TABLE_BODY['F']['USCITE']['title']]
        data.append(bt.TABLE_HEADER[1].format(n=bal_period_end_n.year))
        data.append(bt.TABLE_HEADER[2].format(n_1=bal_period_end_n_1.year))
        data.append(bt.TABLE_BODY['F']['ENTRATE']['title'])
        data.append(bt.TABLE_HEADER[4].format(n=bal_period_end_n.year))
        data.append(bt.TABLE_HEADER[5].format(n_1=bal_period_end_n_1.year))
        data_line.append(data)
        data_line.append(None)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=20, bold=True, wrap=True, fontsize=9, halign='left',
                              border=True)

        # Sezione Cespiti Immobilizzazioni  - Uscite

        if True:
            # Sezione Cespiti Immobilizzazioni  - Uscite - 1
            data_line = [1, 59]
            data = [bt.TABLE_BODY['F']['USCITE'][1]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 59]
            data = [bt.TABLE_BODY['F']['USCITE'][1]['value_n']]
            data.append(bt.TABLE_BODY['F']['USCITE'][1]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione Cespiti Immobilizzazioni  - Uscite - 2
            data_line = [1, 60]
            data = [bt.TABLE_BODY['F']['USCITE'][2]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 60]
            data = [bt.TABLE_BODY['F']['USCITE'][2]['value_n']]
            data.append(bt.TABLE_BODY['F']['USCITE'][2]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione Cespiti Immobilizzazioni  - Uscite - 3
            data_line = [1, 61]
            data = [bt.TABLE_BODY['F']['USCITE'][3]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 61]
            data = [bt.TABLE_BODY['F']['USCITE'][3]['value_n']]
            data.append(bt.TABLE_BODY['F']['USCITE'][3]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione Cespiti Immobilizzazioni  - Uscite - 4
            data_line = [1, 62]
            data = [bt.TABLE_BODY['F']['USCITE'][4]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [2, 62]
            data = [bt.TABLE_BODY['F']['USCITE'][4]['value_n']]
            data.append(bt.TABLE_BODY['F']['USCITE'][4]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione Cespiti Immobilizzazioni  - Uscite - Totalizzatore
            data_line = [1, 64]
            data = ['Totale']
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [2, 64]
            data = ['=SUM(B59:B62)']
            data.append('=SUM(C59:C62)')
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

        # Sezione Cespiti Immobilizzazioni  - Entrate

        if True:
            # Sezione Cespiti Immobilizzazioni  - Entrate - 1
            data_line = [4, 59]
            data = [bt.TABLE_BODY['F']['ENTRATE'][1]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 59]
            data = [bt.TABLE_BODY['F']['ENTRATE'][1]['value_n']]
            data.append(bt.TABLE_BODY['F']['ENTRATE'][1]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione Cespiti Immobilizzazioni  - Entrate - 2
            data_line = [4, 60]
            data = [bt.TABLE_BODY['F']['ENTRATE'][2]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 60]
            data = [bt.TABLE_BODY['F']['ENTRATE'][2]['value_n']]
            data.append(bt.TABLE_BODY['F']['ENTRATE'][2]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione Cespiti Immobilizzazioni  - Entrate - 3
            data_line = [4, 61]
            data = [bt.TABLE_BODY['F']['ENTRATE'][3]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 61]
            data = [bt.TABLE_BODY['F']['ENTRATE'][3]['value_n']]
            data.append(bt.TABLE_BODY['F']['ENTRATE'][3]['value_n_1'])
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione Cespiti Immobilizzazioni  - Entrate - 4
            data_line = [4, 62]
            data = [bt.TABLE_BODY['F']['ENTRATE'][4]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 62]
            data = [bt.TABLE_BODY['F']['ENTRATE'][4]['value_n']]
            # data = [Decimal(0.00)]
            data.append(bt.TABLE_BODY['F']['ENTRATE'][4]['value_n_1'])
            # data.append(Decimal(0.00))
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione Cespiti Immobilizzazioni  - Entrate - 5 Temporary solo 2024 per fusione
            data_line = [4, 63]
            data = [bt.TABLE_BODY['F']['ENTRATE'][5]['DESCRIPTION']]
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                                  border=True)
            data_line = [5, 63]
            data = [bt.TABLE_BODY['F']['ENTRATE'][5]['value_n']]
            # data = [Decimal(0.00)]
            data.append(bt.TABLE_BODY['F']['ENTRATE'][5]['value_n_1'])
            # data.append(Decimal(0.00))
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione Cespiti Immobilizzazioni  - Entrate - Totalizzatore
            data_line = [4, 64]
            data = ['Totale']
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [5, 64]
            data = ['=SUM(E59:E63)']
            data.append('=SUM(F59:F63)')
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione Cespiti Immobilizzazioni  - Imposte
            data_line = [4, 65]
            data = ['Imposte']
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [5, 65]
            data = [Decimal(0.00)]
            data.append(Decimal(0.00))
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

            # Sezione Cespiti Immobilizzazioni  - Avanzo Disavanzo Investimenti
            data_line = [4, 66]
            data = ['Avanzo/disavanzo da entrate e uscite per investimenti e disinvestimenti patrimoniali e finanziamenti']
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)
            data_line = [5, 66]
            data = ['= E64 - B64 - E65']
            data.append('= F64 - C64 -F65')
            data_line.append(data)
            data_line.append(None)
            print(data_line)
            cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                                  border=True)

        # Page Blank small line
        # data_line = [1, 66, ["  "],6]
        # cella = etb.writeline(dataline=data_line, row_height=2, italic=True, wrap=True, fontsize=6, halign='left')

        # Quadro Finale Riassuntivo
        # Headings
        data_line = [5, 67]
        data = [bal_period_end_n.year]
        data.append(bal_period_end_n_1.year)
        data_line.append(data)
        data_line.append(None)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                              border=True)

        # Quadro Finale Riassuntivo
        data_line = [1, 68]
        data = ['Avanzo/disavanzo d’esercizio prima di investimenti e disinvestimenti patrimoniali e finanziamenti']
        data_line.append(data)
        data_line.append(4)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                              border=True)
        data_line = [5, 68]
        data = ['= E56']
        data.append('= F56')
        data_line.append(data)
        data_line.append(None)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                              border=True)
        # Quadro Finale Riassuntivo
        data_line = [1, 69]
        data = ['Avanzo/disavanzo da entrate e uscite per investimenti e disinvestimenti patrimoniali e finanziamenti']
        data_line.append(data)
        data_line.append(4)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                              border=True)
        data_line = [5, 69]
        data = ['= E66']
        data.append('= F66')
        data_line.append(data)
        data_line.append(None)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                              border=True)

        # Quadro Finale Riassuntivo
        data_line = [1, 70]
        data = ['Avanzo/disavanzo complessivo']
        data_line.append(data)
        data_line.append(4)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=22, bold=True, wrap=True, fontsize=10, halign='left',
                              border=True)
        data_line = [5, 70]
        data = ['= E68+E69']
        data.append('= F68+F69')
        data_line.append(data)
        data_line.append(None)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=22, bold=True, wrap=True, fontsize=10, halign='right',
                              border=True)

        # Page Blank small line
        data_line = [1, 71, ["  "],6]
        cella = etb.writeline(dataline=data_line, row_height=2, italic=True, wrap=True, fontsize=6, halign='left')

        # Quadro Finale Riassuntivo - Cassa e banca - Header
        data_line = [1, 72]
        data = ['Cassa e banca']
        data_line.append(data)
        data_line.append(4)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=22, bold=True, wrap=True, fontsize=9, halign='left',
                              border=True)
        data_line = [5, 72]
        data = [bal_period_end_n.year]
        data.append( bal_period_end_n_1.year )
        data_line.append(data)
        data_line.append(None)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=21, bold=True, wrap=True, fontsize=9, halign='right',
                              border=True)

        # Quadro Finale Riassuntivo -
        data_line = [1, 73]
        data = ['Cassa']
        data_line.append(data)
        data_line.append(4)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                              border=True)
        data_line = [5, 73]
        data = [bt.ASSETT['CASSA']['value_n']]
        data.append(bt.ASSETT['CASSA']['value_n_1'])
        data_line.append(data)
        data_line.append(None)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='right',
                              border=True)
        # Quadro Finale Riassuntivo - Banca
        data_line = [1, 74]
        data = ['Depositi bancari e postali']
        data_line.append(data)
        data_line.append(4)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=22, bold=False, wrap=True, fontsize=9, halign='left',
                              border=True)
        data_line = [5, 74]
        data = [bt.ASSETT['BANCA']['value_n']]
        data.append(bt.ASSETT['BANCA']['value_n_1'])
        data_line.append(data)
        data_line.append(None)
        print(data_line)
        cella = etb.writeline(dataline=data_line, row_height=21, bold=False, wrap=True, fontsize=9, halign='right',
                              border=True)



        etb.save()

    # Salva in file compresso i dizionari usati per la compilazione della tabella
    pickle_filename = bt.GNUCASH_FILE.database.split('/')[-1].split('.')[0]+'.pickle'
    global_dicts = {'n':avanzo_conti_periodo_n, 'n_1':avanzo_conti_periodo_n_1, 'bt':bt}
    pickle.dump(global_dicts, open(pickle_filename, "wb"))

    del global_dicts, avanzo_conti_periodo_n, avanzo_conti_periodo_n_1, bt, book




# Main process - formal
if __name__ == '__main__':
    main()
