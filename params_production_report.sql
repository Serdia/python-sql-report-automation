SET NOCOUNT ON;
DECLARE @DateFrom DATE = CAST(DATEADD(mm, DATEDIFF(mm, 0, GETDATE()) - 1, 0) AS DATE),-- first day of previous month 
	    @DateTo DATE = CAST(DATEADD(DAY, - (DAY(GETDATE())), GETDATE()) AS DATE) -- last day of previous month

 -- construct file path to where files should be dropped
DECLARE @FiscalYear INT = CASE 
                            WHEN MONTH(@DateFrom) >= 10 THEN YEAR(@DateFrom) + 1
                            ELSE YEAR(@DateFrom)
                          END
DECLARE @FiscalMonth VARCHAR(100)  = CONCAT( CASE 
										WHEN (MONTH(@DateFrom) + 3) % 12 = 0 THEN 12
										ELSE (MONTH(@DateFrom) + 3) % 12
									END
									,'. '
									,FORMAT(@DateFrom, 'MMMM ')
									,YEAR(@DateFrom)
									)

;WITH cte as (
	SELECT DISTINCT
		CL.CompanyLocationGUID as CarrierLocationGuids,
		l.LineName,
		--CL.LocationName,
		CASE WHEN cl.LocationName = 'Renaissance Reinsurance U.S. Inc on behalf of RenaissanceRe Syndicate 1458' THEN 'Renaissance Reinsurance 1458'
			 WHEN cl.LocationName = 'AXIS Surplus Insurance Company' THEN 'AXIS Surplus Insurance Co'
			 WHEN cl.LocationName = 'Fireman''s Fund Indemnity Corporation' THEN 'Firemans Fund Indemnity Co'
			 WHEN cl.LocationName = 'GuideOne National Insurance Company' THEN 'GuideOne National Insurance Co'
		ELSE CL.LocationName
		END as LocationName,
		co.Location as OfficeName
	FROM tblFin_Invoices INV
	INNER JOIN dbo.tblQuotes q ON INV.QuoteID = q.QuoteID
	INNER JOIN lstLines l ON l.LineGUID = q.LineGuid
	INNER JOIN tblClientOffices co ON co.OfficeGUID = q.QuotingLocationGuid
	INNER JOIN tblFin_InvoiceDetails INVD ON INV.InvoiceNum = INVD.InvoiceNum
	INNER JOIN dbo.tblQuoteDetails ON q.QuoteGUID = dbo.tblQuoteDetails.QuoteGuid
		AND tblQuoteDetails.CompanyLineGuid = INVD.CompanyLineGuid
	LEFT OUTER JOIN dbo.tblCompanyLines ON dbo.tblQuoteDetails.CompanyLineGuid = dbo.tblCompanyLines.CompanyLineGUID
	LEFT OUTER JOIN dbo.tblCompanyLocations AS CL ON dbo.tblCompanyLines.CompanyLocationGUID = CL.CompanyLocationGUID
	LEFT OUTER JOIN tblCompanyLocations ON q.CompanyLocationGuid = tblCompanyLocations.CompanyLocationGUID
	WHERE (INV.Failed = 0)
		AND INVD.ChargeType = 'P'
		AND tblCompanyLocations.CompanyLocationGUID = '1AF7042B-5581-4977-8600-C1BE817C3690' -- Environmental Multi Carrier
		AND cl.LocationName = 'Fireman''s Fund Indemnity Corporation' 

UNION ALL 

	SELECT DISTINCT
		CL.CompanyLocationGUID as CarrierLocationGuids,
		NULL as LineName,
		--CL.LocationName,
		CASE WHEN cl.LocationName = 'Renaissance Reinsurance U.S. Inc on behalf of RenaissanceRe Syndicate 1458' THEN 'Renaissance Reinsurance 1458'
			 WHEN cl.LocationName = 'AXIS Surplus Insurance Company' THEN 'AXIS Surplus Insurance Co'
			 WHEN cl.LocationName = 'Fireman''s Fund Indemnity Corporation' THEN 'Firemans Fund Indemnity Co'
			 WHEN cl.LocationName = 'GuideOne National Insurance Company' THEN 'GuideOne National Insurance Co'
		ELSE CL.LocationName
		END as LocationName,
		co.Location as OfficeName
	FROM tblFin_Invoices INV
	INNER JOIN dbo.tblQuotes q ON INV.QuoteID = q.QuoteID
	INNER JOIN lstLines l ON l.LineGUID = q.LineGuid
	INNER JOIN tblClientOffices co ON co.OfficeGUID = q.QuotingLocationGuid
	INNER JOIN tblFin_InvoiceDetails INVD ON INV.InvoiceNum = INVD.InvoiceNum
	INNER JOIN dbo.tblQuoteDetails ON q.QuoteGUID = dbo.tblQuoteDetails.QuoteGuid
		AND tblQuoteDetails.CompanyLineGuid = INVD.CompanyLineGuid
	LEFT OUTER JOIN dbo.tblCompanyLines ON dbo.tblQuoteDetails.CompanyLineGuid = dbo.tblCompanyLines.CompanyLineGUID
	LEFT OUTER JOIN dbo.tblCompanyLocations AS CL ON dbo.tblCompanyLines.CompanyLocationGUID = CL.CompanyLocationGUID
	LEFT OUTER JOIN tblCompanyLocations ON q.CompanyLocationGuid = tblCompanyLocations.CompanyLocationGUID
	WHERE (INV.Failed = 0)
		AND INVD.ChargeType = 'P'
		AND tblCompanyLocations.CompanyLocationGUID = '1AF7042B-5581-4977-8600-C1BE817C3690' -- Environmental Multi Carrier
		--AND cl.LocationName = 'Fireman''s Fund Indemnity Corporation' 
) 
SELECT 
	@DateFrom as DateFrom 
	,@DateTo as DateTo

	,CONCAT('\\fs01\Align\Accounting\Accounting\Carrier Reporting\', @FiscalYear, ' FY','\DUAL NA\Environmental Subscription\',LocationName, '\', @FiscalMonth) as path
	,OfficeName
	,CONCAT(LocationName, ' Account Current ', FORMAT(@DateFrom, 'MM'), '.', FORMAT(@DateTo, 'yy'),'.xlsx') as FileName
	,CAST(DATEADD(dd,+44,DATEADD(DAY,1,EOMONTH(GETDATE(),-1))) as DATE) as PaymentDue
	,STRING_AGG(CAST(CarrierLocationGuids as VARCHAR(36)), ', ') as CarrierLocationGuids
	,STRING_AGG(CAST(LocationName as VARCHAR(max)), ', ') as CarrierLocationNames
	,CompanyLocationGuids = '1AF7042B-5581-4977-8600-C1BE817C3690' -- Environmental Multi Carrier
	,NULL as LineGuid
	,officeGuid = '778FEC1E-452B-41E2-B630-17CDC5D75D20'  --Dual North America
	,ShowAllOffices = 0
FROM cte
--WHERE LocationName = 'Firemans Fund Indemnity Co'
GROUP BY LocationName,OfficeName