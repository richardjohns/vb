dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
dim bStrm: Set bStrm = createobject("Adodb.Stream")
xHttp.Open "GET", "https://www.rbnz.govt.nz/-/media/ReserveBank/Files/Statistics/tables/b1/hb1-daily.xls", False
xHttp.Send

with bStrm
    .type = 1 '//binary
    .open
    .write xHttp.responseBody
    .savetofile "c:\data\More Data\hb1-daily.xlsx", 2 '//overwrite
end with
