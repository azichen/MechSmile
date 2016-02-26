USE [SMILE_HQ]
GO

/****** Object:  StoredProcedure [dbo].[MMS_R02]    Script Date: 11/14/2012 22:26:55 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/************************************************************
* DataBase : SMILE_HQ
* 程式名稱 : MMS_R02_3
* 程式功能 : 核帳單轉檔彙總表(支票)
* 作    者 : Julia
* 撰寫日期 : 2014/02/10
************************************************************/

CREATE PROCEDURE [dbo].[MMS_R02_3] 
@AA1 VARCHAR(10), --建檔日期(起)
@AA2 VARCHAR(10) --建檔日期(迄)

AS
BEGIN
SET ANSI_WARNINGS OFF;

--DECLARE @AA1 VARCHAR(10), --建檔日期(起)
--@AA2 VARCHAR(10) --建檔日期(迄)

--SET @AA1 = '2013/09/01'
--SET @AA2 = '2013/09/30'

SELECT 
到期日 = C.ChequeExpiryDate,
支票金額 = C.AmountOfCheck,
匯款手續費 = A.RemittanceFee,
銷貨折讓 = A.SalesDiscounts,
銷貨稅額 = A.SalesTax,
其他 = A.Other,
支票號碼 = C.ChequeNo,
銀行帳號 = C.BankAccount,
銀行名稱 = C.PayingBank,
客戶代號 = A.CustomerNo,
發票日期 = B.InvoiceDate,
發票號碼 = B.InvoiceNo,
金額 = B.AmountOfReceive
FROM MMSReceivablesM A LEFT OUTER JOIN MMSReceive B ON A.ReceivablesNo = B.DocNo
LEFT OUTER JOIN MMSCheque C ON A.ReceivablesNo = C.DocNo
WHERE (A.DocDate BETWEEN (CASE @AA1 WHEN '' THEN A.DocDate ELSE @AA1 END) AND (CASE @AA2 WHEN '' THEN A.DocDate  ELSE @AA2 END))
AND A.[Status] = 'Y'


END

GO


