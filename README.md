# Automatización con Macros en VBA

## **¿Qué se hizo?**
Este proyecto creo una automatización inicial y básica con macros en lenguaje VBA de Excel que permitió automatizar múltiples acciones básicas y repetitivas para la creación del reporte general en el área de productos pasivos de una empresa financiera peruana. Estas macros fueron desarrolladas a necesidad del negocio y no se permitió escalar el formato a otro tipo de automatizaciones más complejas.

## **¿Cuál fue el reto?**
El área de productos pasivos tiene en su cartera dos productos financieros: Depósitos a Plazo Fijo y Compensación por Tiempo de Servicios (CTS), contaban con un dashboard creado en Excel y necesitaban automatizar muchas de las acciones que llevaban a cabo por detrás para limpiar y transformar la data de fuentes locales, esto con la finalidad de armar un reporte para gerencia. El reto no fue tanto optimizar la visualización o el reporte como tal, sino reducir el tiempo final de armado. Se redujo de 1 hora a 15 minutos, tiempo que principalmente se demora en juntar la data, ejecutar las macros y realizar algunos cuadres.

## **¿Qué se esperaba conseguir?**
Se quería reducir el tiempo de armado del reporte que permitiría tener más tiempo para decisiones de negocio. Este informe era importante porque permite tener una visualización rápida de como van los dos productos anteriormente mencionados frente a una meta.

## **Programación VBA**
> Preparar Balancines: Entrada y salida de dinero general de la empresa.
```vba
Sub Balancin_DPF_CTS()
  ''ELIMINAR HOJAS BALANCINES JURIDICOS DPF
    Application.DisplayAlerts = False ''Apagar alerta al eliminar
    Sheets("Balancin Dpf").Delete
    Sheets("Balancin DPF - Sin fines").Delete
    Sheets("Balancin DPF - Con fines").Delete
    Application.DisplayAlerts = True ''Prender alerta al eliminar
  ''ELIMINAR COLUMNAS
    Sheets("Balancin DPF Natural").Select
    Columns("A:A").Delete
    Rows("1:4").Delete
  'UNMERGE
    Range("A1").Select
    For I = 1 To 12
      Selection.End(xlToRight).Select
      Selection.UnMerge
    Next I
    Range("A1").Select
  ''BORRAR COLUMNAS EXTRA
    Range("D:G,I:J,M:T,Y:Y,AB:AK,AN:AN,AQ:AQ").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:B").EntireColumn.AutoFit
    Rows("1:1").Delete
    Range("A1").Select
End Sub
```
> Reporte de Cuentas: Todas las cuentas que se crean o cancelan
```vba
Sub Reporte_cuentas_DPF()
  ''BORRAR HOJAS
  Application.DisplayAlerts = False
  Sheets("Todo DPF").Delete
  Sheets("DPF Juridica - Sin fines").Delete
  Sheets("DPF Juridica - Con fines").Delete
  Application.DisplayAlerts = True
  Sheets("DPF Natural").Select
  ''COLUMNAS REALES
  Columns("A:A").Delete
  Rows("1:4").Delete
  Range("B:B,L:L,N:O,Q:Q,S:AB,AD:AD,AG:AH,AJ:AK").Select
  Selection.Delete Shift:=xlToLeft
  Range("A1").Select
End Sub
```

## **¿Siguientes pasos?**
Reunirse con el negocio para poder solicitar que se conecten las bases locales en una sola tabla en SQL y poder exportar la data limpia.
