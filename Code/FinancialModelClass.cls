VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FinancialModelClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Outputs
Public costDifference_R As Double
Public systemCost_R As Double
Public NPV_R As Double
Public IRR As Double
Public LCOH_RperKWh As Double
Public systemEfficiency As Double
Public solarFraction As Double

'Inputs
Dim dRate As Double
Dim exchangeRate_R2EUR As Double
Dim costFuel_RperkWh As Double
Dim priceIncreasePAn As Double

'In object variable
Dim differenceInRunningCosts As Double
Dim serviceLife As Double


Sub constructor(discountRate As Double, exchangeRate_rand2Euro As Double, costFuel_RandperkWh As Double, priceIncreasePerAn As Double)
    dRate = discountRate
    exchangeRate_R2EUR = exchangeRate_rand2Euro
    costFuel_RperkWh = costFuel_RandperkWh
    priceIncreasePAn = priceIncreasePerAn
    
    serviceLife = 20
    costDifference_R = 0
End Sub

Sub CostDifference_SPHvsHeater(E_heaterNoSPH_kWh As Double, E_heaterWithSPH_kWh As Double)
    
    costDifference_R = (E_heaterNoSPH_kWh - E_heaterWithSPH_kWh) * costFuel_RperkWh
    
End Sub

Sub calcIRR_and_NPV()
    Dim cashFlows() As Double
    Dim n As Integer
    
    ReDim cashFlows(serviceLife)
    
    cashFlows(0) = -systemCost_R
    
    For n = 1 To serviceLife
        cashFlows(n) = (1 + priceIncreasePAn ^ n) * costDifference_R
    Next n
    
    IRR = WorksheetFunction.IRR(cashFlows)
    NPV_R = NPV(dRate, cashFlows)
End Sub

Sub calcNPV()
    Dim n As Integer
    
    NPV_R = -systemCost_R
    
    For n = 1 To serviceLife
        NPV_R = NPV_R + ((1 + (priceIncreasePAn) ^ n) * costDifference_R) / ((1 + dRate) ^ n)
    Next n
End Sub

Sub calcLCOH_R(E_solarOneYear_kWh As Double)
    Dim numerator As Double
    Dim denominator As Double
    Dim n As Integer
    
    numerator = systemCost_R
    denominator = 0
    
    For n = 1 To serviceLife
        denominator = denominator + E_solarOneYear_kWh / ((1 + dRate) ^ n)
    Next n
    
    LCOH_RperKWh = numerator / denominator
End Sub

Sub calcSystemCost_R(grossColArea As Double)
    Dim specificSystemCost_EUR As Double
    
    specificSystemCost_EUR = -0.41 * grossColArea + 770.34
    
    systemCost_R = exchangeRate_R2EUR * (specificSystemCost_EUR * grossColArea)
    systemCost_R = exchangeRate_R2EUR * (503 * grossColArea)
End Sub

Sub calcSolarFraction(E_solar As Double, E_auxilary As Double)
    Dim E_total As Double
    
    E_total = E_solar + E_auxilary
    
    solarFraction = E_solar / E_total
End Sub

Sub calcSystemEfficiency(E_solar As Double, I_solar As Double)
    systemEfficiency = E_solar / I_solar
End Sub
