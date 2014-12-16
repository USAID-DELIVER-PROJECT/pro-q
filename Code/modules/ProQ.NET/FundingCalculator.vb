Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("FundingCalculator_NET.FundingCalculator")> Public Class FundingCalculator
	' FundingCalculator
	'
	' This module will Handle the Allocation of Funding and calculate
	' the available funds in the Fundings, Funding Sources, Use Fundings
	' Costs, Quantifications, and costs.
	
	' 14-June-2002 jleiner
	'
	' - WORRIED ABOUT THE SCOPE, MUST ONLY BE CALLED IN FRONT END? TO KNOW Agg
	
	Function GetAllocated_UseFunding(ByRef objAggregation As Aggregation, ByRef objUF As UseFunding) As Double
		' Comments  : Loop through the funding records for the different
		'           : costs and return the the sum of those for this Use Funding
		' Parametes : ObjUF the Use Funding in Question
		' Returns   : Double - The amount of Funds allocated
		' Created   : 14-June-2002 jleiner
		'------------------------------------------------------------------------
		
		Dim dblFunds As Double
		'Dim g_objAggregation As Aggregation
		
		Dim objQ As Quantification
		Dim objCost As Cost
		Dim objFunding As Funding
		Dim i As Short
		Dim intCurrentCostObject As Short
		
		
		dblFunds = 0
		i = 1
		
		'Move Through all the cost Objects for each Quantification in the Agg
		'Then move through each Funding Record
		For	Each objQ In objAggregation.m_colQuantifications
			
			'Trap the Current Cost so to set it back if in middle of process
			intCurrentCostObject = objQ.getCurrentCost.GetCategory()
			
			'For loop to set the current Cost records and then go through
			'it Funding Records
			'------------------------------------------------------------
			For i = 1 To 3
				
				' Set the Current Cost
				objQ.SetCurrentCost(i)
				objCost = objQ.getCurrentCost
				
				'Check each funding for a fit. If the Funding Record is of the
				'Specified Same Funding Source and Use or (no Use if UF.use=*)
				'------------------------------------------------------------
				For	Each objFunding In objCost.m_colFundings
					If objFunding.GetFundingSourceID = objUF.GetFundingSourceID Then
						If objUF.GetUseID = "*" Then
							dblFunds = dblFunds + objFunding.GetValue()
						ElseIf objQ.GetUse.GetID = objUF.GetUseID Then 
							dblFunds = dblFunds + objFunding.GetValue()
						End If
					End If
				Next objFunding
			Next i
			
			objQ.SetCurrentCost(intCurrentCostObject)
			
		Next objQ
		
		GetAllocated_UseFunding = dblFunds
		
	End Function
	
	Function GetAllocated_FundingSource(ByRef objAggregation As Aggregation, ByRef objFS As FundingSource) As Double
		' Comments  : Loop through the funding records for the different
		'           : costs and return the the sum of those for this Funding Source
		' Parametes : ObjFS - the Funding Source in Question
		' Returns   : Double - The amount of Funds allocated
		' Created   : 14-June-2002 jleiner
		'------------------------------------------------------------------------
		
		Dim dblFunds As Double
		
		Dim objQ As Quantification
		Dim objCost As Cost
		Dim objFunding As Funding
		Dim i As Short
		Dim intCurrentCostObject As Short
		
		dblFunds = 0
		i = 1
		
		'Move Through all the cost Objects for each Quantification in the Agg
		'Then move through each Funding Record
		For	Each objQ In objAggregation.m_colQuantifications
			
			'Trap the Current Cost so to set it back if in middle of process
			intCurrentCostObject = objQ.getCurrentCost.GetCategory()
			
			'For loop to set the current Cost records and then go through
			'it Funding Records
			'------------------------------------------------------------
			For i = 1 To 3
				
				' Set the Current Cost
				objQ.SetCurrentCost(i)
				objCost = objQ.getCurrentCost
				
				'Check each funding for a fit. If the Funding Record is of the
				'Specified Same Funding Source and Use or (no Use if UF.use=*)
				'------------------------------------------------------------
				For	Each objFunding In objCost.m_colFundings
					If objFunding.GetFundingSourceID = objFS.GetID Then
						dblFunds = dblFunds + objFunding.GetValue()
					End If
				Next objFunding
			Next i
			
			objQ.SetCurrentCost(intCurrentCostObject)
			
		Next objQ
		
		GetAllocated_FundingSource = dblFunds
		
	End Function
End Class