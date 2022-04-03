Public Sub LiquidaBoletaCob(CreditoID As Long, OpcionCobranza As Integer, CantCuotas As Integer, PorcDeduccion As Double, FechaLiq As Date)
    On Error GoTo errLiquidaBoleta
    
    Dim linea As Recordset, RstCuadroCred As Recordset, RstTempLiquidacion As Recordset, RstTempLiquAgrupado As Recordset
    Dim UltCuota As Integer, AuxCuota As Integer, AuxVto As Date
    Dim PrimerVencimiento As Date
    Dim FechaImpago As Date
    Dim Atrasado, Bandera, PrimeraCuotaImpaga As Boolean
    Dim CapitalCobrado, InteresCobrado, Resarcitorios, Ajuste, GastoAdmin, SaldoCapital As Currency
    Dim DiasAjuste, DiasAjusteTotal As Long
    Dim ImpuestoII As Single
    Dim FechaDeLiquidacion As Date
    
    FechaDeLiquidacion = FechaLiq
    
    'Tomo datos del crédito y de la línea para poder calcular el cuadro de marcha de aquel entonces
    Set linea = CreditosDb.OpenRecordset("SELECT Creditos.CreditoID, Creditos.CreditoFechaProxVto, LineasCreditos.LineaCreditoDescrip, PeriodosAmortizaciones.PeriodoAmortizacionMeses, " & _
        "LineasCreditos.LineaPeriodoGraciaTope, Creditos.CreditoCapitalLiq, Creditos.CreditoCapitalLiq, Creditos.CreditoCuoPag, LineasCreditos.LineaCreditoID, " & _
        "LineasCreditos.TipoVtoId, LineasCreditos.LineaSellado, LineasCreditos.LineaGastoOriginacion, LineasCreditos.LineaSeguroVida, Creditos.CreditoCuoConv, " & _
        "LineasCreditos.LineaFdoGarantia, LineasCreditos.LineaIVASobreInt, LineasCreditos.LineaGastosAdmin, LineasCreditos.LineaIVAGeneral, Creditos.CreditoFechaLiquidacion " & _
        "FROM (LineasCreditos INNER JOIN Creditos ON LineasCreditos.LineaCreditoID = Creditos.LineaCreditoID) INNER JOIN PeriodosAmortizaciones ON LineasCreditos.PeriodoAmortizacionID = PeriodosAmortizaciones.PeriodoAmortizacionID " & _
        "WHERE Creditos.CreditoID = " + Trim(CreditoID), dbOpenDynaset, dbSeeChanges)  'La query se llama : InicioCargo"
    
    'Calculo el primer vencimiento por que el Att (CreditoFechaPrimCuota) no es confiable
    Bandera = True
    If Not IsNull(linea!TipoVtoId) Then
        PrimerVencimiento = Det1erVTO(linea!LineaCreditoID, linea!CreditoFechaLiquidacion)
    Else
        MsgBox "Falta el tipo de vencimiento de la línea.", vbCritical, "Atención!!!"
        Bandera = False
    End If
    
    If IsNull(linea!LineaPeriodoGraciaTope) Then
        MsgBox "Falta el período de gracia.", vbCritical
        Bandera = False
    End If
    
    If IsNull(linea!PeriodoAmortizacionMeses) Then
        MsgBox "Falta el período de amortización (en meses).", vbCritical, "Atención!!!"
        Bandera = False
    End If
    
    If IsNull(linea!LineaGastosAdmin) Then
        MsgBox "Falta el porcentaje de gastos administrativos a cobrar.", vbCritical, "Atención!!!"
        Bandera = False
    End If
    
    If IsNull(linea!LineaSeguroVida) Then
        MsgBox "Falta el porcentaje del seguro de vida a cobrar.", vbCritical, "Atención!!!"
        Bandera = False
    End If
    
    If IsNull(linea!LineaFdoGarantia) Then
        MsgBox "Falta el porcentaje de fondo de garantía a cobrar.", vbCritical, "Atención!!!"
        Bandera = False
    End If
    
    If IsNull(linea!LineaIVASobreInt) Then
        MsgBox "Falta el porcentaje de IVA sobre intereses a cobrar.", vbCritical, "Atención!!!"
        Bandera = False
    End If
    
    If IsNull(linea!LineaIVAGeneral) Then
        MsgBox "Falta el porcentaje de IVA general a cobrar.", vbCritical, "Atención!!!"
        Bandera = False
    End If
    If Bandera = False Then
        Exit Sub
    End If
    
    Set RstCuadroCred = CreditosDb.OpenRecordset("SELECT CreditoID, CuadroNumCuota, ConceptoCuotaId, CuadroFVto, CuadroImporteDeb, CuadroImporteCred, CuadroImporteSaldo " & _
        "FROM CuadroCreditos " & _
        "WHERE CreditoID=" & CreditoID & " AND CuadroImporteSaldo>0 " & _
        "ORDER BY CuadroNumCuota, ConceptoCuotaId", dbOpenDynaset, dbSeeChanges)
    
    Set RstTempLiquidacion = CreditosDb.OpenRecordset("TempGestionCobranza")
    If Not RstCuadroCred.EOF Then
        Select Case OpcionCobranza
        Case 1  ' Vencidas
            RstCuadroCred.MoveFirst
            AuxCuota = RstCuadroCred!CuadroNumCuota + CantCuotas - 1
            With RstCuadroCred
                Do While RstCuadroCred!CuadroNumCuota <= AuxCuota
                    RstTempLiquidacion.AddNew
                    RstTempLiquidacion!CreditoID = RstCuadroCred!CreditoID
                    RstTempLiquidacion!NumCuota = RstCuadroCred!CuadroNumCuota
                    RstTempLiquidacion!VtoCuota = DetVTOCuota(RstCuadroCred!CreditoID, RstCuadroCred!CuadroNumCuota)
                    RstTempLiquidacion!ConceptoCuotaID = RstCuadroCred!ConceptoCuotaID
                    RstTempLiquidacion!ImporteConcepto = Redondeo2(RstCuadroCred!CuadroImporteSaldo)
                    RstTempLiquidacion!FechaLiquidacion = FechaDeLiquidacion
                    RstTempLiquidacion.Update
                    RstCuadroCred.MoveNext
                Loop
            End With
        Case 2 ' Debengadas
            RstCuadroCred.MoveFirst
            AuxCuota = 0
            With RstCuadroCred
                Do Until RstCuadroCred!CuadroNumCuota = AuxCuota
                    AuxVto = DetVTOCuota(RstCuadroCred!CreditoID, RstCuadroCred!CuadroNumCuota)
                    If AuxVto > FechaDeLiquidacion Then
                        AuxCuota = RstCuadroCred!CuadroNumCuota
                    End If
                    RstCuadroCred.MoveNext
                Loop
                RstCuadroCred.MoveFirst
                Do Until RstCuadroCred.EOF
                    If RstCuadroCred!CuadroNumCuota = AuxCuota Then
                        RstTempLiquidacion.AddNew
                        RstTempLiquidacion!CreditoID = RstCuadroCred!CreditoID
                        RstTempLiquidacion!NumCuota = RstCuadroCred!CuadroNumCuota
                        RstTempLiquidacion!VtoCuota = DetVTOCuota(RstCuadroCred!CreditoID, RstCuadroCred!CuadroNumCuota)
                        RstTempLiquidacion!ConceptoCuotaID = RstCuadroCred!ConceptoCuotaID
                        RstTempLiquidacion!ImporteConcepto = Redondeo2(RstCuadroCred!CuadroImporteSaldo)
                        RstTempLiquidacion!FechaLiquidacion = FechaDeLiquidacion
                        RstTempLiquidacion.Update
                    End If
                    RstCuadroCred.MoveNext
                Loop
            End With
            
        Case 3 ' Vencidas y devengadas
            RstCuadroCred.MoveFirst
            AuxCuota = 0
            With RstCuadroCred
                Do Until RstCuadroCred!CuadroNumCuota = AuxCuota
                    AuxVto = DetVTOCuota(RstCuadroCred!CreditoID, RstCuadroCred!CuadroNumCuota)
                    If AuxVto > FechaDeLiquidacion Then
                        AuxCuota = RstCuadroCred!CuadroNumCuota
                    End If
                    RstCuadroCred.MoveNext
                Loop
                RstCuadroCred.MoveFirst
                Do While RstCuadroCred!CuadroNumCuota <= AuxCuota
                    RstTempLiquidacion.AddNew
                    RstTempLiquidacion!CreditoID = RstCuadroCred!CreditoID
                    RstTempLiquidacion!NumCuota = RstCuadroCred!CuadroNumCuota
                    RstTempLiquidacion!VtoCuota = DetVTOCuota(RstCuadroCred!CreditoID, RstCuadroCred!CuadroNumCuota)
                    RstTempLiquidacion!ConceptoCuotaID = RstCuadroCred!ConceptoCuotaID
                    RstTempLiquidacion!ImporteConcepto = Redondeo2(RstCuadroCred!CuadroImporteSaldo)
                    RstTempLiquidacion!FechaLiquidacion = FechaDeLiquidacion
                    RstTempLiquidacion.Update
                    RstCuadroCred.MoveNext
                Loop
            End With
            
        Case 4 ' Cancelación
            RstCuadroCred.MoveFirst
            With RstCuadroCred
                'AuxCuota = RstCuadroCred!CuadroNumCuota
                Do Until RstCuadroCred!CuadroNumCuota = AuxCuota
                    AuxVto = DetVTOCuota(RstCuadroCred!CreditoID, RstCuadroCred!CuadroNumCuota)
                    If AuxVto >= FechaDeLiquidacion Then
                        AuxCuota = RstCuadroCred!CuadroNumCuota
                    End If
                    RstCuadroCred.MoveNext
                    If RstCuadroCred.EOF Then
                        If AuxCuota = 0 Then
                            AuxCuota = 10000
                        Else
                            AuxCuota = 1
                        End If
                        Exit Do
                    End If
                Loop
                RstCuadroCred.MoveFirst
                Do Until RstCuadroCred.EOF
                    If RstCuadroCred!CuadroNumCuota > AuxCuota Then
                        If RstCuadroCred!ConceptoCuotaID = 1 Then
                            RstTempLiquidacion.AddNew
                            RstTempLiquidacion!CreditoID = RstCuadroCred!CreditoID
                            RstTempLiquidacion!NumCuota = RstCuadroCred!CuadroNumCuota
                            RstTempLiquidacion!VtoCuota = DetVTOCuota(RstCuadroCred!CreditoID, RstCuadroCred!CuadroNumCuota)
                            RstTempLiquidacion!ConceptoCuotaID = RstCuadroCred!ConceptoCuotaID
                            RstTempLiquidacion!ImporteConcepto = Redondeo2(RstCuadroCred!CuadroImporteSaldo)
                            RstTempLiquidacion!FechaLiquidacion = FechaDeLiquidacion
                            RstTempLiquidacion.Update
                        End If
                    Else
                        RstTempLiquidacion.AddNew
                        RstTempLiquidacion!CreditoID = RstCuadroCred!CreditoID
                        RstTempLiquidacion!NumCuota = RstCuadroCred!CuadroNumCuota
                        RstTempLiquidacion!VtoCuota = DetVTOCuota(RstCuadroCred!CreditoID, RstCuadroCred!CuadroNumCuota)
                        RstTempLiquidacion!ConceptoCuotaID = RstCuadroCred!ConceptoCuotaID
                        RstTempLiquidacion!ImporteConcepto = Redondeo2(RstCuadroCred!CuadroImporteSaldo)
                        RstTempLiquidacion!FechaLiquidacion = FechaDeLiquidacion
                        RstTempLiquidacion.Update
                    End If
                    RstCuadroCred.MoveNext
                Loop
            End With
        End Select
    End If
    
    Set RstTempLiquAgrupado = openrs("SELECT NumCuota FROM TempGestionCobranza GROUP BY NumCuota ORDER BY NumCuota")
    
    If RstTempLiquAgrupado.NoMatch = False Then
        RstTempLiquAgrupado.MoveFirst
        With RstTempLiquAgrupado
            'AuxCuota = 0
            Do Until RstTempLiquAgrupado.EOF
                If RstTempLiquAgrupado!NumCuota < AuxCuota + 1 Then
                    'cargo Resarcitorio
                    If (DetResarcitorio(CreditoID, RstTempLiquAgrupado!NumCuota, FechaDeLiquidacion)) > 0 Then
                        RstTempLiquidacion.AddNew
                        RstTempLiquidacion!CreditoID = CreditoID
                        RstTempLiquidacion!NumCuota = RstTempLiquAgrupado!NumCuota
                        RstTempLiquidacion!VtoCuota = DetVTOCuota(RstTempLiquidacion!CreditoID, RstTempLiquAgrupado!NumCuota)
                        RstTempLiquidacion!ConceptoCuotaID = 8
                        RstTempLiquidacion!ImporteConcepto = Redondeo2(DetResarcitorio(CreditoID, RstTempLiquAgrupado!NumCuota, FechaDeLiquidacion))
                        RstTempLiquidacion!DatoAdicionalText = DateDiff("D", RstTempLiquidacion!VtoCuota, FechaDeLiquidacion) & " Días"
                        RstTempLiquidacion!DatoAdicionalNum = DateDiff("D", RstTempLiquidacion!VtoCuota, FechaDeLiquidacion)
                        RstTempLiquidacion!FechaLiquidacion = FechaDeLiquidacion
                        RstTempLiquidacion.Update
                    End If
                    'cargo Bonificacion
                    If (DetResarcitorio(CreditoID, RstTempLiquAgrupado!NumCuota, Now()) * PorcDeduccion / 100) > 0 Then
                        RstTempLiquidacion.AddNew
                        RstTempLiquidacion!CreditoID = CreditoID
                        RstTempLiquidacion!NumCuota = RstTempLiquAgrupado!NumCuota
                        RstTempLiquidacion!VtoCuota = DetVTOCuota(RstTempLiquidacion!CreditoID, RstTempLiquAgrupado!NumCuota)
                        RstTempLiquidacion!ConceptoCuotaID = 12
                        RstTempLiquidacion!ImporteConcepto = Redondeo2(DetResarcitorio(CreditoID, RstTempLiquAgrupado!NumCuota, FechaDeLiquidacion) * PorcDeduccion / 100 * -1)
                        RstTempLiquidacion!DatoAdicionalText = PorcDeduccion & " %"
                        RstTempLiquidacion!DatoAdicionalNum = PorcDeduccion
                        RstTempLiquidacion!FechaLiquidacion = FechaDeLiquidacion
                        RstTempLiquidacion.Update
                    End If
                    'Cargo iva Resarcitotio
                    If (((DetResarcitorio(CreditoID, RstTempLiquAgrupado!NumCuota, FechaDeLiquidacion)) - (DetResarcitorio(CreditoID, RstTempLiquAgrupado!NumCuota, FechaDeLiquidacion) * PorcDeduccion / 100)) * (linea!LineaIVASobreInt / 100)) > 0 Then
                        RstTempLiquidacion.AddNew
                        RstTempLiquidacion!CreditoID = CreditoID
                        RstTempLiquidacion!NumCuota = RstTempLiquAgrupado!NumCuota
                        RstTempLiquidacion!VtoCuota = DetVTOCuota(RstTempLiquidacion!CreditoID, RstTempLiquAgrupado!NumCuota)
                        RstTempLiquidacion!ConceptoCuotaID = 33
                        RstTempLiquidacion!ImporteConcepto = Redondeo2(((DetResarcitorio(CreditoID, RstTempLiquAgrupado!NumCuota, FechaDeLiquidacion)) - (DetResarcitorio(CreditoID, RstTempLiquAgrupado!NumCuota, FechaDeLiquidacion) * PorcDeduccion / 100)) * (linea!LineaIVASobreInt / 100))
                        RstTempLiquidacion!FechaLiquidacion = FechaDeLiquidacion
                        RstTempLiquidacion.Update
                    End If
                    'End If
                End If
                'AuxCuota = RstTempLiquidacion!CuadroNumCuota
                RstTempLiquAgrupado.MoveNext
            Loop
        End With
    End If
    
    Exit Sub
errLiquidaBoleta:
    MsgBox Err.description
End Sub