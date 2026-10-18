let
    Origen = Excel.CurrentWorkbook(){[Name="Tabla11"]}[Content],
    #"Tipo cambiado" = Table.TransformColumnTypes(Origen,{{"Columna1", type text}, {"Columna2", type text}, {"Columna3", type text}, {"Columna4", type any}, {"Columna5", type any}, {"Columna6", type any}, {"Columna7", type any}, {"Columna8", type text}, {"Columna9", type any}, {"Columna10", type any}, {"Columna11", type any}, {"Columna12", type text}, {"Columna13", type any}, {"Columna14", type any}, {"Columna15", type any}, {"Columna16", type any}, {"Columna17", type any}}),
    #"Columnas con nombre cambiado2" = Table.RenameColumns(#"Tipo cambiado",{{"Columna1", " Cuenta"}, {"Columna3", " Descripcion"}, {"Columna9", "Saldo anterior"}, {"Columna11", "Debitos "}, {"Columna13", "Creditos "}, {"Columna16", "Saldo"}, {"Columna2", "Column2"}, {"Columna4", "Column4"}, {"Columna5", "Column5"}, {"Columna6", "Column6"}, {"Columna7", "Column7"}, {"Columna8", "Column8"}, {"Columna10", "Column10"}, {"Columna12", "Column12"}, {"Columna14", "Column14"}, {"Columna15", "Column15"}, {"Columna17", "Column17"}}),
    #"Filas superiores quitadas" = Table.Skip(#"Columnas con nombre cambiado2",7),
    #"Tipo cambiado1" = Table.TransformColumnTypes(#"Filas superiores quitadas",{{" Cuenta", type text}, {"Column2", type text}, {" Descripcion", type text}, {"Column4", type any}, {"Column5", type any}, {"Column6", type any}, {"Column7", type any}, {"Column8", type text}, {"Saldo anterior", type any}, {"Column10", type any}, {"Debitos ", type any}, {"Column12", type text}, {"Creditos ", type any}, {"Column14", type any}, {"Column15", type any}, {"Saldo", type any}, {"Column17", type any}}),
    #"Filas en blanco eliminadas" = Table.SelectRows(#"Tipo cambiado1", each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}))),
    #"Columnas con nombre cambiado" = Table.RenameColumns(#"Filas en blanco eliminadas",{{" Cuenta", "PUC"}, {"Saldo anterior", "Saldo inicial"}, {"Saldo", "Saldo final"}, {"Column2", "Column2"}}),
    #"Columnas reordenadas" = Table.ReorderColumns(#"Columnas con nombre cambiado",{"PUC", "Column2"," Descripcion", "Saldo inicial", "Debitos ", "Creditos ", "Saldo final"}),
    #"Tipo cambiado2" = Table.TransformColumnTypes(#"Columnas reordenadas",{{"PUC", Int64.Type}}),
    #"Errores quitados" = Table.RemoveRowsWithErrors(#"Tipo cambiado2", {"PUC"}),
    #"Tipo cambiado3" = Table.TransformColumnTypes(#"Errores quitados",{{"Saldo inicial", Currency.Type}, {"Debitos ", Currency.Type}, {"Creditos ", Currency.Type}, {"Saldo final", Currency.Type}}),
    #"Errores quitados1" = Table.RemoveRowsWithErrors(#"Tipo cambiado3", {"Saldo inicial"}),
    #"Filas filtradas" = Table.SelectRows(#"Errores quitados1", each [#" Descripcion"] <> "afdsaf"),
    #"Personalizada agregada1" = Table.AddColumn(#"Filas filtradas", "Personalizado", each if [Saldo inicial] is null then [#" Descripcion"] else ""),
    #"Columnas reordenadas1" = Table.ReorderColumns(#"Personalizada agregada1",{"PUC", " Descripcion", "Personalizado", "Saldo inicial", "Debitos ", "Creditos ", "Saldo final"}),

    // --- Aquí desplazamos la columna ---
ColName = "Personalizado",
Names = Table.ColumnNames(#"Columnas reordenadas1"),
i = List.PositionOf(Names, ColName),
ListaOriginal = Table.Column(#"Columnas reordenadas1", ColName),
ListaShift = List.Range(ListaOriginal, 1) & {null},
Lists = Table.ToColumns(#"Columnas reordenadas1"),
ListsSinCol = List.RemoveRange(Lists, i, 1),
ListsConShift = List.InsertRange(ListsSinCol, i, {ListaShift}),
Result = Table.FromColumns(ListsConShift, Names),
    #"Columnas combinadas" = Table.CombineColumns(Result,{" Descripcion", "Personalizado"},Combiner.CombineTextByDelimiter(" ", QuoteStyle.None),"Combinada"),
    #"Filas filtradas1" = Table.SelectRows(#"Columnas combinadas", each ([Saldo inicial] <> null)),
    #"Columnas quitadas" = Table.RemoveColumns(#"Filas filtradas1",{"Column4", "Column5", "Column6", "Column7", "Column8", "Column10", "Column12", "Column14", "Column15", "Column17"}),
    #"Poner En Mayúsculas Cada Palabra" = Table.TransformColumns(#"Columnas quitadas",{{"Combinada", Text.Proper, type text}}),
    #"Personalizada agregada" = Table.AddColumn(#"Poner En Mayúsculas Cada Palabra", "Tercero", each if [Column2] is null then "" else [Combinada]),
    #"Columnas reordenadas2" = Table.ReorderColumns(#"Personalizada agregada",{"PUC", "Column2", "Tercero", "Combinada", "Saldo inicial", "Debitos ", "Creditos ", "Saldo final"}),
    #"Columnas quitadas1" = Table.RemoveColumns(#"Columnas reordenadas2",{"Column2"}),
    #"Tipo cambiado4" = Table.TransformColumnTypes(#"Columnas quitadas1",{{"PUC", type text}}),
    #"Personalizada agregada2" = Table.AddColumn(#"Tipo cambiado4", "Personalizado", each [PUC]&"01"),
    #"Columnas reordenadas3" = Table.ReorderColumns(#"Personalizada agregada2",{"PUC", "Personalizado", "Tercero", "Combinada", "Saldo inicial", "Debitos ", "Creditos ", "Saldo final"}),
    #"Rellenar hacia abajo" = Table.FillDown(#"Columnas reordenadas3",{"Personalizado"}),
    #"Personalizada agregada3" = Table.AddColumn(#"Rellenar hacia abajo", "Personalizado.1", each if [Tercero] ="" then [PUC] else [Personalizado]),
    #"Columnas reordenadas4" = Table.ReorderColumns(#"Personalizada agregada3",{"PUC", "Personalizado", "Personalizado.1", "Tercero", "Combinada", "Saldo inicial", "Debitos ", "Creditos ", "Saldo final"}),
    #"Columnas quitadas2" = Table.RemoveColumns(#"Columnas reordenadas4",{"PUC", "Personalizado"}),
    #"Columnas con nombre cambiado1" = Table.RenameColumns(#"Columnas quitadas2",{{"Personalizado.1", "PUC"}, {"Combinada", "Descripción"}}),
    #"Columnas reordenadas5" = Table.ReorderColumns(#"Columnas con nombre cambiado1",{"PUC", "Descripción", "Tercero", "Saldo inicial", "Debitos ", "Creditos ", "Saldo final"}),
    #"Personalizada agregada4" = Table.AddColumn(#"Columnas reordenadas5", "Cuenta", each if Text.Start([PUC], 1) = "1" then "Activo"
else if Text.Start([PUC], 1) = "2" then "Pasivo"
else if Text.Start([PUC], 1) = "3" then "Patrimonio"
else if Text.Start([PUC], 1) = "4" then "Ingreso"
else if Text.Start([PUC], 1) = "5" then "Gasto"
else "Costo"),
    #"Columnas reordenadas6" = Table.ReorderColumns(#"Personalizada agregada4",{"Cuenta", "PUC", "Descripción", "Tercero", "Saldo inicial", "Debitos ", "Creditos ", "Saldo final"}),
    #"Personalizada agregada5" = Table.AddColumn(#"Columnas reordenadas6", "Largo", each Text.Length([PUC])),
    #"Columnas reordenadas7" = Table.ReorderColumns(#"Personalizada agregada5",{"Largo", "Cuenta", "PUC", "Descripción", "Tercero", "Saldo inicial", "Debitos ", "Creditos ", "Saldo final"}),
    #"Personalizada agregada6" = Table.AddColumn(#"Columnas reordenadas7", "Dato Financiero", each if Text.Start([PUC], 1) = "1" then [Saldo final]
else if Text.Start([PUC], 1) = "2" then [Saldo final]*-1
else if Text.Start([PUC], 1) = "3" then [Saldo final]*-1
else if Text.Start([PUC], 1) = "4" then [#"Creditos "]-[#"Debitos "]
else [#"Debitos "]-[#"Creditos "]),
    #"Tipo cambiado5" = Table.TransformColumnTypes(#"Personalizada agregada6",{{"Dato Financiero", Currency.Type}})
in
    #"Tipo cambiado5"
