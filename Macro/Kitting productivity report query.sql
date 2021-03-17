/*import VAS Billing as [VAS Billing] and Work Order Summary as [WO Summary 1]*/

SELECT DISTINCTROW Record.[WO#], [VAS Billing].[Kitted Qty], Sum(Record.TotalManHr) AS [Sum Of TotalManHr], Round([VAS Billing].[Kitted Qty]/Sum(Record.TotalManhr),0) AS [PcsPerManhr], Round([VAS Billing].[Kitted Qty]*16.09/[VAS Billing].[Total Kitting Cost],0) AS [ExpectedPcsPerManhrs],[WO Summary 1].[FG SKU], [WO Summary 1].[JOB DESCRIPTION],Round([PcsPerManhr]-[ExpectedPcsPerManhrs],0) AS [Diff in productivity], Round([VAS Billing].[Kitted Qty]/[Diff in productivity],0) AS [Saved Manhr], [Saved Manhr]*16.09 AS [Margin],[Margin]/[VAS Billing].[Kitted Qty] AS [MarginPerFGUnit]

FROM ((Record INNER JOIN [WO Summary 1] ON Record.[WO#] = [WO Summary 1].[WO #])INNER JOIN [VAS Billing] ON Record.[WO#] = [VAS Billing].[WO #])
GROUP BY Record.[WO#], [VAS Billing].[Kitted Qty],[VAS Billing].[Kitted Qty]*16.09/[VAS Billing].[Total Kitting Cost],[WO Summary 1].[FG SKU], [WO Summary 1].[JOB DESCRIPTION]