@ECHO OFF
SETLOCAL

SET _source="C:\Bioptics"

SET _dest="J:\Silva's Lab\P30 Core Center\Faxitron Backup\Bioptics"

SET _what=/MIR
:: /COPYALL :: COPY ALL file info
:: /B :: copy files in Backup mode.
:: /SEC :: copy files with SECurity
:: /MIR :: MIRror a directory tree

ROBOCOPY %_source% %_dest% %_what% 