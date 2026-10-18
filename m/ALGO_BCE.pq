let
    Origen = Excel.CurrentWorkbook(){[Name="Tabla9"]}[Content],
    #"Columnas quitadas" = Table.RemoveColumns(Origen,{"Nivel", "Transaccional", "Identificación", "Sucursal"}),
    #"Poner En Mayúsculas Cada Palabra" = Table.TransformColumns(#"Columnas quitadas",{{"Nombre tercero", Text.Proper, type text}}),
    #"Tipo cambiado" = Table.TransformColumnTypes(#"Poner En Mayúsculas Cada Palabra",{{"Saldo Inicial", Currency.Type}, {"Movimiento Débito", Currency.Type}, {"Movimiento Crédito", Currency.Type}, {"Saldo final", Currency.Type}, {"Código cuenta contable", type text}}),
    #"Personalizada agregada" = Table.AddColumn(#"Tipo cambiado", "Largo", each Text.Length([Código cuenta contable])),
    #"Columnas reordenadas" = Table.ReorderColumns(#"Personalizada agregada",{"Año", "Largo", "Mes", "Código cuenta contable", "Nombre Cuenta contable", "Nombre tercero", "Saldo Inicial", "Movimiento Débito", "Movimiento Crédito", "Saldo final"}),
    #"Personalizada agregada1" = Table.AddColumn(#"Columnas reordenadas", "Cuenta", each if Text.Start([Código cuenta contable],1) = "1" then "Activo" 
else if Text.Start([Código cuenta contable],1) = "2" then "Pasivo"
else if Text.Start([Código cuenta contable],1) = "3" then "Patrimonio"
else ""),
    #"Columnas reordenadas1" = Table.ReorderColumns(#"Personalizada agregada1",{"Año", "Mes", "Largo", "Cuenta", "Código cuenta contable", "Nombre Cuenta contable", "Nombre tercero", "Saldo Inicial", "Movimiento Débito", "Movimiento Crédito", "Saldo final"}),
    #"Filas filtradas" = Table.SelectRows(#"Columnas reordenadas1", each ([Cuenta] <> "")),
    #"Personalizada agregada2" = Table.AddColumn(#"Filas filtradas", "Dato Financiero", each if Text.Start([Código cuenta contable],1) = "1" then [Saldo final] 
else [Saldo final] * -1),
    #"Tipo cambiado1" = Table.TransformColumnTypes(#"Personalizada agregada2",{{"Dato Financiero", Currency.Type}}),
    #"Personalizada agregada3" = Table.AddColumn(#"Tipo cambiado1", "Puc & Descripción", each [Código cuenta contable] &" "& [Nombre Cuenta contable]),
    #"Columnas reordenadas2" = Table.ReorderColumns(#"Personalizada agregada3",{"Año", "Mes", "Largo", "Cuenta", "Código cuenta contable", "Nombre Cuenta contable", "Nombre tercero", "Puc & Descripción", "Saldo Inicial", "Movimiento Débito", "Movimiento Crédito", "Saldo final", "Dato Financiero"})
in
    #"Columnas reordenadas2"
