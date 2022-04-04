public sub AgregarRetencion(byval Concepto as integer, byval importe As Currency, byref RetencionAC As Currency)
    With openrs("RetencionesTempo")
        .AddNew
        !ConceptoCuotaID = Concepto 
        !TempoImporte = importe
        !CodigoConcepto = 1
        !SucursalID = TempVars!SucursalID
        !CreditoID = 0
        !PcIP = TempVars!IPPc
        RetencionAC = RetencionAC + !TempoImporte
        .Update
        .close
    end with
end sub