# Gnucash-piecash Bilancio per APS Italiane

Gli script Python sviluppati in questo progetto servono per estrarre le informazioni contabili inserite in un programma di contabilità chiamato GnuCash, catalogarle secondo i criteri stabiliti dalla normativa articolo 13 del Codice del Terzo settore (dlgs 117/2017):

> Gli enti del Terzo settore devono redigere il bilancio di esercizio formato dallo stato
patrimoniale, dal rendiconto gestionale, con l'indicazione dei proventi e degli oneri
dell'ente, e dalla relazione di missione che illustra le poste di bilancio, l'andamento
economico e gestionale dell'ente e le modalità di perseguimento delle finalità
statutarie.

> Il bilancio degli enti del Terzo settore con ricavi, rendite, proventi o entrate
comunque denominate inferiori a 220.000 euro può essere redatto nella forma del
*rendiconto per cassa*.

> Il bilancio di cui ai commi 1 e 2 deve essere redatto in conformità alla modulistica
definita con decreto del Ministro del Lavoro e delle Politiche sociali, sentito il
Consiglio Nazionale del Terzo settore.

Il progetto **piecash_APS_balance** consente quindi di "mappare" i conti utilizzati con Gnucash per poterli aggregare nelle voci previste dal "rendiconto per cassa" previsto come adempimento di bilancio da presentare al **RUNTS** ossia **R**egistro **U**nico **N**azionale del **T**erzo **S**ettore.

## Che cos'è un'Associazione di Promozione Sociale APS?

L’**A**ssociazione di **P**romozione **S**ociale (**APS**) è stata introdotta nell’ordinamento italiano dalla legge 383/2000.
In base al Codice del Terzo Settore è un Ente del Terzo Settore e pertanto deve presentarne le caratteristiche essenziali, quindi l’assenza di fini di lucro e lo svolgimento di un’attività d’interesse generale.
In quanto Associazione di Promozione Sociale, invece deve assumere la forma dell’Associazione ed essere composta da non meno di sette persone fisiche o tre Associazioni di Promozione Sociale.

Possono ammettere come soci anche altri Enti del Terzo Settore o senza scopo di lucro ma questi non devono superare il 50% delle Associazioni di Promozione Sociale socie. Sono previste eccezioni per gli enti di natura sportiva.
Possono avvalersi del lavoro dipendente o autonomo, anche dei propri associati, se necessario ai fini dello svolgimento dell’attività di interesse generale o al raggiungimento delle proprie finalità, ma il numero dei lavoratori non può superare il 50% dei volontari o il 5% degli associati.

Non possono essere riconosciute come Associazioni di Promozione Sociale, Associazioni o circoli privati che pongono qualsiasi tipi di discriminazione all’accesso, incluse le condizioni economiche e patrimoniali, o richiedano la partecipazione a quote di natura patrimoniale.

## Che cos'è GnuCash

GnuCash è un programma finanziario e di contabilità adatto all'utilizzo in ambito famigliare o in una piccola impresa, rilasciato gratuitamente con licenza [GNU](https://www.gnu.org/) GPL e disponibile per GNU/Linux, BSD, Solaris, Mac OS X e Microsoft Windows.

Progettato per essere di semplice utilizzo, ma comunque potente e flessibile, GnuCash permette di tenere traccia dei conti bancari, delle azioni, delle entrate e delle uscite. Intuitivo nell'utilizzo come il registro del libretto degli assegni, si basa sui principi fondamentali della contabilità per garantire l'equilibrio dei saldi e l'accuratezza dei resoconti.

GnuCash è sviluppato, mantenuto, documentato e tradotto interamente da volontari. Vuoi offire il tuo contributo? Abbiamo alcuni [suggerimenti](https://wiki.gnucash.org/wiki/Contributing_to_GnuCash). Aiuta a tradurre GnuCash nella tua lingua su [Weblate](https://hosted.weblate.org/engage/gnucash/).

### Caratteristiche principali

- Contabilità a partita doppia
- Conti per azioni, obbligazioni e fondi comuni
- Contabilità di piccole imprese
- Report, Grafici
- Importazione di dati QIF/OFX/HBCI, ricerca delle corrispondenze tra le transazioni
- Transazioni pianificate
- Calcoli finanziari

# piecash

Join the chat at
- https://gitter.im/sdementen/piecash 
- https://readthedocs.org/projects/piecash/badge/?version=master 
- https://coveralls.io/repos/sdementen/piecash/badge.svg?branch=master&service=github

**Piecash** provides a simple and pythonic interface to GnuCash files stored in SQL (sqlite3, Postgres and MySQL).

Documentation:	http://piecash.readthedocs.org.

* Gitter:	https://gitter.im/sdementen/piecash

* Github:	https://github.com/sdementen/piecash

* PyPI:	https://pypi.python.org/pypi/piecash

It is a pure python package, tested on python 3.6 to 3.9, that can be used as an alternative to:

- the official python bindings (as long as no advanced book modifications and/or engine calculations are needed). This is specially useful on Windows where the official python bindings may be tricky to install or if you want to work with python 3.

- XML parsing/reading of XML GnuCash files if you prefer python over XML/XLST manipulations.

- piecash test suite runs successfully on Windows and Linux on the three supported SQL backends (sqlite3, Postgres and MySQL). piecash has also been successfully run on Android (sqlite3 backend) thanks to Kivy buildozer and python-for-android.

It allows you to:

- open existing GnuCash documents and access all objects within
- modify objects or add new objects (accounts, transactions, prices, ...)
- create new GnuCash documents from scratch