# xlsx-mapper
Command Line interface that maps Metadata from Excel-Template to EditShare Metadata-Fields 

## Install
Download & run xlsx-mapper.msi

## How to use
Starte Eingabeaufforderung / Windows Powershell
Empfehlung: [Windows Terminal](https://learn.microsoft.com/en-us/windows/terminal/install)

#### Zeigt hilfe an: <code>xlsx-mapper -h</code>

#### Einmalige konfiguration eines EditShare users: <code>xlsx-mapper -c</code>
Username & password werden in einer Datei gespeichert, daher muss <code>xlsx-mapper -c</code> nur einmalig ausgeführt werden. <code>xlsx-mapper -c</code> kann aber erneut ausgeführt werden z.B. zum umloggen eines Users.

#### Mapping: <code>xlsx-mapper -m dateipfad</code>
