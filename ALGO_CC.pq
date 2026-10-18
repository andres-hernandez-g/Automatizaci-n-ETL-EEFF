let
    Origen = Excel.CurrentWorkbook(){[Name="Tabla12"]}[Content],
    #"Tipo cambiado" = Table.TransformColumnTypes(Origen,{{"Año", Int64.Type}, {"Mes", type text}, {"Nivel", type text}, {"Transaccional", type text}, {"Código cuenta contable", Int64.Type}, {"Nombre cuenta contable", type text}, {"Código centro costo", type text}, {"Nombre centro costo", type text}, {"Saldo inicial", Int64.Type}, {"Movimiento débito", type number}, {"Movimiento crédito", type number}, {"Saldo final", type number}}),
    #"Columnas quitadas" = Table.RemoveColumns(#"Tipo cambiado",{"Nivel", "Transaccional"}),
    #"Dividir columna por delimitador" = Table.SplitColumn(#"Columnas quitadas", "Código centro costo", Splitter.SplitTextByDelimiter("-", QuoteStyle.Csv), {"Código centro costo.1", "Código centro costo.2"}),
    #"Tipo cambiado1" = Table.TransformColumnTypes(#"Dividir columna por delimitador",{{"Código centro costo.1", Int64.Type}, {"Código centro costo.2", Int64.Type}}),
    #"Columnas combinadas" = Table.CombineColumns(Table.TransformColumnTypes(#"Tipo cambiado1", {{"Código centro costo.2", type text}}, "es-419"),{"Código centro costo.2", "Nombre centro costo"},Combiner.CombineTextByDelimiter(" ", QuoteStyle.None),"Combinada"),
    #"Columnas con nombre cambiado" = Table.RenameColumns(#"Columnas combinadas",{{"Combinada", "Subcuenta"}}),
    #"Columnas quitadas1" = Table.RemoveColumns(#"Columnas con nombre cambiado",{"Código centro costo.1"}),
    #"Tipo cambiado2" = Table.TransformColumnTypes(#"Columnas quitadas1",{{"Código cuenta contable", type text}, {"Saldo inicial", Currency.Type}, {"Movimiento débito", Currency.Type}, {"Movimiento crédito", Currency.Type}, {"Saldo final", Currency.Type}}),
    #"Columnas reordenadas" = Table.ReorderColumns(#"Tipo cambiado2",{"Año", "Mes", "Subcuenta", "Código cuenta contable", "Nombre cuenta contable", "Saldo inicial", "Movimiento débito", "Movimiento crédito", "Saldo final"}),
    #"Personalizada agregada" = Table.AddColumn(#"Columnas reordenadas", "Largo", each Text.Length([Código cuenta contable])),
    #"Columnas reordenadas1" = Table.ReorderColumns(#"Personalizada agregada",{"Año", "Mes", "Largo", "Subcuenta", "Código cuenta contable", "Nombre cuenta contable", "Saldo inicial", "Movimiento débito", "Movimiento crédito", "Saldo final"}),
    #"Personalizada agregada1" = Table.AddColumn(#"Columnas reordenadas1", "Cuenta", each if Text.Start([Código cuenta contable],1) = "4" then "Ingreso"
else if Text.Start([Código cuenta contable],1) = "5" then "Gasto"
else "Costo"),
    #"Columnas reordenadas2" = Table.ReorderColumns(#"Personalizada agregada1",{"Año", "Mes", "Largo", "Cuenta", "Subcuenta", "Código cuenta contable", "Nombre cuenta contable", "Saldo inicial", "Movimiento débito", "Movimiento crédito", "Saldo final"}),
    #"Personalizada agregada2" = Table.AddColumn(#"Columnas reordenadas2", "Dato Financiero", each if Text.Start([Código cuenta contable],1) = "4" then [Movimiento crédito]-[Movimiento débito]
else [Movimiento débito]-[Movimiento crédito]),
    #"Tipo cambiado3" = Table.TransformColumnTypes(#"Personalizada agregada2",{{"Dato Financiero", Currency.Type}}),
    #"Personalizada agregada3" = Table.AddColumn(#"Tipo cambiado3", "Puc & Descripción", each [Código cuenta contable] &" "&[Nombre cuenta contable]),
    #"Columnas reordenadas3" = Table.ReorderColumns(#"Personalizada agregada3",{"Año", "Mes", "Largo", "Cuenta", "Subcuenta", "Código cuenta contable", "Nombre cuenta contable", "Puc & Descripción", "Saldo inicial", "Movimiento débito", "Movimiento crédito", "Saldo final", "Dato Financiero"})
in
    #"Columnas reordenadas3"