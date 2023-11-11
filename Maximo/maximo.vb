Dim tempSQL as string
 	tempSQL = "let" & vbCrlf 
tempSQL = tempSQL & "	 " & vbcrlf
tempSQL = tempSQL & "    Source = Sql.Database(""ssrs-db.vnnapplications.com:1492"", ""ProjectDB""), " & vbcrlf
tempSQL = tempSQL & "    MAXIMO_TICKET = Source{[Schema=""MAXIMO"",Item=""TICKET""]}[Data], " & vbcrlf
tempSQL = tempSQL & "    #""Remove Unused Columns"" = Table.SelectColumns(MAXIMO_TICKET,{ ""REPORTEDPRIORITY"", ""CLASS"", ""TICKETID"",  ""DESCRIPTION"", ""STATUS"", ""TARGETSTART"", ""TARGETFINISH"", ""SC_ANALDUEDATE"", ""SC_OWNER""}), " & vbcrlf
tempSQL = tempSQL & "    #""Merge Columns"" = Table.CombineColumns(Table.TransformColumnTypes(#""Remove Unused Columns"", {{""TARGETFINISH"", type text}, {""SC_ANALDUEDATE"", type text}}, ""en-US""),{""TARGETFINISH"", ""SC_ANALDUEDATE""}, Combiner.CombineTextByDelimiter("""", QuoteStyle.None),""DUE""), " & vbcrlf
tempSQL = tempSQL & "    #""Change Type"" = Table.TransformColumnTypes(#""Merge Columns"",{{""DUE"", type datetime}}), " & vbcrlf
tempSQL = tempSQL & "    #""Sort Rows"" = Table.Sort(#""Change Type"",{{""DUE"", Order.Ascending}}), " & vbcrlf
tempSQL = tempSQL & "    #""Filter Out Closed and Canceled Items"" = Table.SelectRows(#""Sort Rows"", each ([STATUS] <> ""CANCEL"" and [STATUS] <> ""CLOSED"" and [STATUS] <> ""CLOSEDCR"" and [STATUS] <> ""CLOSEDSR"")), " & vbcrlf
tempSQL = tempSQL & "    #""Filter On Names"" = Table.SelectRows(#""Filter Out Closed and Canceled Items"", each  " & vbcrlf
tempSQL = tempSQL & "	    ( " & vbcrlf
tempSQL = tempSQL & "					 " & vbcrlf
tempSQL = tempSQL & "					[SC_OWNER] =           ""VBALAKRI"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""EJBEAN"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""BHBENDER"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2CTBLAC"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2JBLAZE"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""AMBOERNE"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""MJBOUCHE"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""PGBRADAT"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2JRBRAN"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2JGBRIN"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""JTCORBET"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2BCUSIC"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""ddarr"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2ELDICK"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2NERTLE"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2MFRENC"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""pgibbs"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2LGIHYE"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2NGRAFT"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2RGRIME"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2DEHARM"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2BHIRMA"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""CMHOWARD"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""AHUSSEIN"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2LKATZE"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2PGMANS"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""DCMCCORM"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2RHNORV"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""DAOGLESB"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2BDORTI"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""RFPILUSO"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""PDPOTTER"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""RLRUNYON"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2MSSHOO"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""DESPIELM"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2DSSPIN"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2DSREDN"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2KPVIKA"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2BWERNE"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""X2TAWOOL"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""KAYENNER"" " & vbcrlf
tempSQL = tempSQL & "               or [SC_OWNER] = ""SMGALE"" " & vbcrlf
tempSQL = tempSQL & "			   )), " & vbcrlf
tempSQL = tempSQL & "    #""Filter Dates"" = Table.SelectRows(#""Filter On Names"", each Date.IsInNextNWeeks([DUE], 4)), " & vbcrlf
tempSQL = tempSQL & "    #""Add Notes Column"" = Table.AddColumn(#""Filter Dates"", ""Notes"", each if [CLASS] <> null then """" else null), " & vbcrlf
tempSQL = tempSQL & "    #""Add Index Column"" = Table.AddIndexColumn(#""Add Notes Column"", ""Index"", 0, 1), " & vbcrlf
tempSQL = tempSQL & "    #""Reorder Columns"" = Table.ReorderColumns(#""Add Index Column"",{""Index"", ""REPORTEDPRIORITY"", ""CLASS"", ""TICKETID"", ""DESCRIPTION"", ""STATUS"", ""TARGETSTART"", ""DUE"", ""SC_OWNER"", ""Notes""}), " & vbcrlf
tempSQL = tempSQL & "    #""Rename Columns"" = Table.RenameColumns(#""Reorder Columns"",{{""CLASS"", ""Type""}, {""REPORTEDPRIORITY"", ""P""}, {""TARGETSTART"", ""Start""}, {""SC_OWNER"", ""Owner""},{""Index"",""I""}}) " & vbcrlf
tempSQL = tempSQL & "in " & vbcrlf
tempSQL = tempSQL & "    #""Rename Columns"" " & vbcrlf