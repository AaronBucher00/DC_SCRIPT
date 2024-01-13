für DE siehe unten.

EN:

Problem definition
The project work IFC to Minergie-Excel focussed on the use of the Minergie-Excel tool 
for the verification of summer thermal insulation in accordance with SIA180. This 
This tool enables the assessment of conformity with the standard on the basis of room information.
Despite its effectiveness, the application requires manual transfer of room and project information, which 
project information, which makes handling time-consuming. In established processes 
processes, this method is therefore only used for rooms that the project manager classifies as necessary 
or critical by the project manager, which is a potential source of error.
The restriction to selected rooms harbours the risk that the least favourable room may not be 
the most unfavourable room for verification, which may affect the reliability of the results. 
may affect the reliability of the results.


Solution concept
To overcome the challenges posed by the manual transfer of room and project information 
project information, a solution was developed that includes both a technical solution and an optimised process. 
optimised process. As part of this approach, a PythonScript was created, 
that uses an existing IFC model as a source. In addition, the script guides the 
project manager through the required information inputs using a user-friendly 
user-friendly interface. Subsequently, on the basis of the queries made 
The Minergie Excel is then completed for all rooms in the project based on the queries made. A 
subsequent project-specific control of the expenditure by the project manager is necessary for the 
is essential for use as a standard-compliant verification for summer thermal insulation in accordance with SIA 180.

In order to ensure that the models used can be optimally utilised, care was taken to 
to work exclusively with the Base Quantities in IFC. This approach ensures that 
the programme can be used universally and is not tied to specific IFC properties. 
By automating the process and integrating an existing IFC model, the source of error that is 
minimises the source of errors that can arise when manually selecting critical spaces.


functions and Classes:

main:
  glb_exp_data = [ ]: global variable: Stores the evaluated data
  start_main_func (): calls individual functions

import_ifc:
  load_source_ifc (): allows the user to select an IFC file
  read_source_ifc (): reads and stores necessary data from the IFC file

export_ifc_data:
  export_ifc_data (): saves IFC data as an Excel file

minergie_excel_editor:
  Minergie_Excel_Editor : class for gathering user input (project-specific information)
  save_values (): saves values to multiple Excel files




DE:

Problemstellung
In der Projektarbeit IFC to Minergie-Excel wurde sich mit der Verwendung des Minergie-Excel-Tools 
für den Nachweis des sommerlichen  Wärmeschutzes gemäß SIA180 befasst. Dieses 
Instrument ermöglicht die Beurteilung der Normkonformität anhand von Rauminformationen.
Trotz seiner Effektivität erfordert die Anwendung  manuelle Übertragungen von Raum- und 
Projektinformationen, was die Handhabung zeitaufwändig gestaltet. In etablierten Prozessen 
wird dieses Verfahren daher lediglich für Räume eingesetzt, die vom Projektleiter als notwendig 
oder kritisch eingestuft werden, was potenziell eine Fehlerquelle darstellt.
Die Beschränkung auf ausgewählte Räume birgt das Risiko, dass möglicherweise nicht der 
ungünstigste Raum für den Nachweis erfasst wird, was die Zuverlässigkeit der Ergebnisse 
beeinträchtigen kann.


Lösungskonzept
Um die Herausforderungen der manuellen Übertragung von Raum- und Projektinformationen 
zu bewältigen, wurde eine Lösung entwickelt, das sowohl einen technische Lösung als auch einen 
optimierten Prozess beinhaltet. Im Rahmen dieses Ansatzes wurde ein Python-Script erstellt, 
das als Quelle ein vorhandenes IFC-Modell nutzt. Zusätzlich führt das Skript den 
Projektleiter durch die erforderlichen Informationsinputs mithilfe einer 
benutzerfreundlichen Oberfläche. Anschliessend wird auf Grundlage der getätigten 
Abfragen, das Minergie-Excel, für sämtliche im Projekt vorhandenen Räume, ausgefüllt. Eine 
nachträgliche projektspezifische Kontrolle der Ausgaben durch den Projektleiter, ist für den 
Einsatz als normkonformen Nachweis für den sommerlichen Wärmeschutz nach SIA 180 unumgänglich.

Um sicherzustellen, dass die verwendeten Modelle optimal genutzt werden können, wurde darauf 
geachtet, ausschließlich mit den Base Quantities in IFC zu arbeiten. Dieser Ansatz gewährleistet, dass 
das Programm allgemeingültig eingesetzt werden kann und nicht an spezifische IFC-Propertys gebunden ist. 
Durch die Automatisierung des Prozesses und die Integration eines vorhandenen IFC-Modells wird 
die Fehlerquelle, die bei der manuellen Auswahl kritischer Räume entstehen kann, minimiert.


Funktionen und Klassen:

main:
  glb_exp_data = [ ] globale Variable: Speichern der ausgewerteten Daten
  start_main_func (): Aufrufen der einzelnen Funktionen

import_ifc:
  load_source_ifc (): User IFC-Datei auswählen lassen
  read_source_ifc (): benötigte Daten aus IFC-Datei auslesen und speichern

export_ifc_data:
  export_ifc_data (): IFC-Daten als Excel speichern

minergie_excel_editor:
  Minergie_Excel_Editor : Klasse: Input von User abholen (Projektspezifische Info)
  save_values (): Speichern von mehreren Excel-Files
