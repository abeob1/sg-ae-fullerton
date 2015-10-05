
CALL AE_SP001_AnaplanTrailBalExtraction('2015-01-01','2015-03-31');
DROP  PROCEDURE AE_SP001_AnaplanTrailBalExtraction;

CREATE PROCEDURE AE_SP001_AnaplanTrailBalExtraction
(
 IN DateFrom DATE,
IN DateTo DATE
)

AS
CompanyCode varchar(100);

BEGIN

SELECT TT0."U_ANAPLANCODE" into CompanyCode  FROM "FHG_TEST_24042015"."@AI_TB01_COMPANYDATA" TT0 
WHERE TT0."Name"=(SELECT CURRENT_SCHEMA FROM DUMMY);

CREATE COLUMN TABLE "TBEXTRACT" (ReportingUnit varchar(100)
								,ICPartnerCode Varchar(200),AcctCode varchar(200)
								,Movement varchar(10),Amount Decimal);
								
CREATE COLUMN TABLE "TBFINAL" (ReportingDate DATE,ReportingUnit varchar(100)
								,ICPartnerCode Varchar(200),AcctCode varchar(200)
								,Movement varchar(10),Amount Decimal);								
--================================ Balance Amount ==================================================================
INSERT INTO "TBEXTRACT"(
								
SELECT
T0."ShortName",
T1."U_ANAPLANICCODE" AS "Anaplan IC Code"
,T0."Account"
,'' AS "Movement"
,SUM(T0."Debit")-SUM(T0."Credit") AS "Amount"

FROM 
"JDT1" T0
LEFT JOIN "OCRD" T1 ON T0."ShortName"=T1."CardCode"
RIGHT JOIN "OACT" T2 ON T0."Account"=T2."AcctCode"

WHERE T0."RefDate"<=:DateTo
AND LEFT (T0."ShortName",1)='Z'
GROUP BY T0."Account",T2."AcctCode",T1."U_ANAPLANICCODE",T0."ShortName");


INSERT INTO "TBEXTRACT"(
								
SELECT
T0."ShortName",
T1."U_ANAPLANICCODE" AS "Anaplan IC Code"
,T0."Account"
,'' AS "Movement"
,SUM(T0."Debit")-SUM(T0."Credit") AS "Amount"

FROM 
"JDT1" T0
LEFT JOIN "OCRD" T1 ON T0."ShortName"=T1."CardCode"
RIGHT JOIN "OACT" T2 ON T0."Account"=T2."AcctCode"

WHERE T0."RefDate"<=:DateTo
AND LEFT (T0."ShortName",1)<>'Z'
GROUP BY T0."Account",T2."AcctCode",T1."U_ANAPLANICCODE",T0."ShortName");


--=================================== Opening Balance ==========================================================================
INSERT INTO "TBEXTRACT"(
								
SELECT
T0."ShortName",
T1."U_ANAPLANICCODE" AS "Anaplan IC Code"
,T0."Account"
,'M00' AS "Movement"
,SUM(T0."Debit")-SUM(T0."Credit") AS "Amount"

FROM 
"JDT1" T0
LEFT JOIN "OCRD" T1 ON T0."ShortName"=T1."CardCode"
RIGHT JOIN "OACT" T2 ON T0."Account"=T2."AcctCode"

WHERE T0."RefDate"<:DateFrom
AND LEFT (T0."ShortName",1)='Z'
GROUP BY T0."Account",T2."AcctCode",T1."U_ANAPLANICCODE",T0."ShortName");


INSERT INTO "TBEXTRACT"(
								
SELECT
T0."ShortName",
T1."U_ANAPLANICCODE" AS "Anaplan IC Code"
,T0."Account"
,'M00' AS "Movement"
,SUM(T0."Debit")-SUM(T0."Credit") AS "Amount"

FROM 
"JDT1" T0
LEFT JOIN "OCRD" T1 ON T0."ShortName"=T1."CardCode"
RIGHT JOIN "OACT" T2 ON T0."Account"=T2."AcctCode"

WHERE T0."RefDate"<:DateFrom 
AND LEFT (T0."ShortName",1)<>'Z'
GROUP BY T0."Account",T2."AcctCode",T1."U_ANAPLANICCODE",T0."ShortName");


--================================== Debit Amount ===============================================================

INSERT INTO "TBEXTRACT"(
								
SELECT
T0."ShortName",
T1."U_ANAPLANICCODE" AS "Anaplan IC Code"
,T0."Account"
,'M05' AS "Movement"
,SUM(T0."Debit")  AS "Amount"

FROM 
"JDT1" T0
LEFT JOIN "OCRD" T1 ON T0."ShortName"=T1."CardCode"
RIGHT JOIN "OACT" T2 ON T0."Account"=T2."AcctCode"

WHERE T0."RefDate" BETWEEN :DateFrom AND :DateTo
AND LEFT (T0."ShortName",1)='Z'
GROUP BY T0."Account",T2."AcctCode",T1."U_ANAPLANICCODE",T0."ShortName");


INSERT INTO "TBEXTRACT"(
								
SELECT
T0."ShortName",
T1."U_ANAPLANICCODE" AS "Anaplan IC Code"
,T0."Account"
,'M05' AS "Movement"
,SUM(T0."Debit") AS "Amount"

FROM 
"JDT1" T0
LEFT JOIN "OCRD" T1 ON T0."ShortName"=T1."CardCode"
RIGHT JOIN "OACT" T2 ON T0."Account"=T2."AcctCode"

WHERE T0."RefDate" BETWEEN :DateFrom AND :DateTo 
AND LEFT (T0."ShortName",1)<>'Z'
GROUP BY T0."Account",T2."AcctCode",T1."U_ANAPLANICCODE",T0."ShortName");


--=================================== Credit Amout ==========================================================================

INSERT INTO "TBEXTRACT"(
								
SELECT
T0."ShortName",
T1."U_ANAPLANICCODE" AS "Anaplan IC Code"
,T0."Account"
,'M06' AS "Movement"
,SUM(T0."Credit") AS "Amount"

FROM 
"JDT1" T0
LEFT JOIN "OCRD" T1 ON T0."ShortName"=T1."CardCode"
RIGHT JOIN "OACT" T2 ON T0."Account"=T2."AcctCode"

WHERE T2."AcctCode"='1-11010-01'AND T0."RefDate" BETWEEN :DateFrom AND :DateTo
AND LEFT (T0."ShortName",1)='Z'
GROUP BY T0."Account",T2."AcctCode",T1."U_ANAPLANICCODE",T0."ShortName");


INSERT INTO "TBEXTRACT"(
								
SELECT
T0."ShortName",
T1."U_ANAPLANICCODE" AS "Anaplan IC Code"
,T0."Account"
,'M06' AS "Movement"
,SUM(T0."Credit") AS "Amount"

FROM 
"JDT1" T0
LEFT JOIN "OCRD" T1 ON T0."ShortName"=T1."CardCode"
RIGHT JOIN "OACT" T2 ON T0."Account"=T2."AcctCode"

WHERE T0."RefDate" BETWEEN :DateFrom AND :DateTo 
AND LEFT (T0."ShortName",1)<>'Z'
GROUP BY T0."Account",T2."AcctCode",T1."U_ANAPLANICCODE",T0."ShortName");

--=============================================================================================================
INSERT INTO "TBFINAL"(
SELECT 
TO_DATE(:DateTo,'YYYY-MM-DD')AS "Repporting Date"
,:CompanyCode AS "Anaplan Entity Code"
--,MAX(T0."REPORTINGUNIT")
,T0."ICPARTNERCODE"
,T0."ACCTCODE"
,T0."MOVEMENT"
,SUM(T0."AMOUNT")

 FROM "TBEXTRACT" T0 GROUP BY T0."ICPARTNERCODE",T0."ACCTCODE"
,T0."MOVEMENT");
 
 
SELECT * FROM "TBFINAL" T0 ORDER BY T0."ACCTCODE";

DROP TABLE "TBEXTRACT";
DROP TABLE "TBFINAL";

END;