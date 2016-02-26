USE [SMILE_HQ]
GO

/****** Object:  StoredProcedure [dbo].[MMS_R02]    Script Date: 11/14/2012 22:26:55 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

/************************************************************
* DataBase : SMILE_HQ
* �{���W�� : MMS_R02_1
* �{���\�� : �ֱb�����ɷJ�`��(�{��)
* �@    �� : Julia
* ���g��� : 2014/02/10
************************************************************/

CREATE PROCEDURE [dbo].[MMS_R02_1] 
@AA1 VARCHAR(10), --���ɤ��(�_)
@AA2 VARCHAR(10) --���ɤ��(��)

AS
BEGIN
SET ANSI_WARNINGS OFF;

--DECLARE @AA1 VARCHAR(10), --���ɤ��(�_)
--@AA2 VARCHAR(10) --���ɤ��(��)

--SET @AA1 = '2013/09/01'
--SET @AA2 = '2013/09/30'

SELECT 
���ڤ� = A.DateOfReceipt,
�w�s�{�� = A.Cash,
�P�f���� = A.SalesDiscounts,
�P�f�|�B = A.SalesTax,
��L = A.Other,
�Ȥ�N�� = A.CustomerNo,
�o����� = B.InvoiceDate,
�o�����X = B.InvoiceNo,
���B = B.AmountOfReceive
FROM MMSReceivablesM A LEFT OUTER JOIN MMSReceive B ON A.ReceivablesNo = B.DocNo
WHERE (A.DocDate BETWEEN (CASE @AA1 WHEN '' THEN A.DocDate ELSE @AA1 END) AND (CASE @AA2 WHEN '' THEN A.DocDate  ELSE @AA2 END))
AND A.[Status] = 'Y'


END

GO


