-- Copyright 2006-2020 ConnectCode Pte Ltd. All Rights Reserved.
-- This source code is protected by Copyright Laws. You are only allowed to modify
-- and include the source in your application if you have purchased a Distribution License.
-- ===================================================================================
-- ConnectCode Barcode TSQL UDF for SQL Server
--
-- The formulas in this file can be used for creating barcodes in Microsoft SQL SErver.

CREATE OR ALTER FUNCTION dbo.UPCAUNICHR(@DataNum  int )
  RETURNS nchar
  AS
  BEGIN
      DECLARE @ReturnString nchar;
	  SET @ReturnString='';
	  
	  IF @DataNum>=158 and @DataNum<=167  
		   SET @ReturnString = char(@DataNum);
	  ELSE IF @DataNum>=192 and @DataNum<=216
		   SET @ReturnString = char(@DataNum);
      ELSE
		   SET @ReturnString = char(@DataNum);
    
      RETURN(@ReturnString);
  END
GO
CREATE OR ALTER FUNCTION dbo.generateUPCACheckDigit(@Data nvarchar(255))
  RETURNS nvarchar(255)
  AS
  BEGIN
   DECLARE @Datalength int=LEN(@Data);
   DECLARE @Parity int=0; 
   DECLARE @Sumvalue int=0;
   DECLARE @Resultvalue int=-1;
   DECLARE @StrResult nvarchar(30)='';
   DECLARE @Barcodechar nvarchar(1)='';
   DECLARE @Barcodevalue int=0;
   DECLARE @X int=1;
   
   While @X <=@Datalength
   BEGIN
    SET @Barcodechar=SUBSTRING(@Data,@X,1);
	SET @Barcodevalue=ASCII(@Barcodechar);
	SET @Barcodevalue=@Barcodevalue - 48;
	IF @X % 2=1 
		SET @Sumvalue = @Sumvalue + (3*@Barcodevalue);
	ELSE
		SET @Sumvalue = @Sumvalue + @Barcodevalue;	
    SET @X=@X+1; 
   END

   SET @Resultvalue = @Sumvalue % 10;
   IF @Resultvalue = 0 
	SET @Resultvalue = 0;
   ELSE
	SET @Resultvalue = 10 - @Resultvalue;

   SET @Resultvalue = @Resultvalue+48;	
   
   SET @StrResult=@StrResult + dbo.UNICHR(@Resultvalue);

   RETURN @StrResult;
  END
  GO


CREATE OR ALTER FUNCTION dbo.EncodeUPCA( @InputData nvarchar(255), @Hr int )
  RETURNS nvarchar(255)
  AS
  BEGIN    
	  DECLARE @TempChar nvarchar(1);
	  DECLARE @Transformchar nvarchar(1);
	  DECLARE @Weight int=1;
	  DECLARE @Transformvalue int=0;
	  DECLARE @SumValue int=0;
	  DECLARE @OneValue int=0;
	  DECLARE @ChkValue int=0;
	  DECLARE @Lendiff int=0;
	  DECLARE @CD char;
      DECLARE @FilteredData nvarchar(255)='';
      DECLARE @ReturnData nvarchar(255)=''; 
      DECLARE @ResultData nvarchar(255)=''; 
	  DECLARE @FilteredLength int=LEN(@InputData);	  
   	  DECLARE @Paritybit    int=0;
   	  DECLARE @Firstdigit   int=0; 
   	  DECLARE @Datalength   int=0;
   	  DECLARE @Transformdataleft nvarchar(255)='';
   	  DECLARE @Transformdataright nvarchar(255)='';

	  DECLARE @Counter int;
	  DECLARE @X int;

	  IF @FilteredLength <= 0 
	  BEGIN
	      SET @ReturnData = ''; 
		  RETURN(@ReturnData);
	  END
      SET @Counter = 1;
	  While @Counter<=@FilteredLength
	  BEGIN
			SET @TempChar = SUBSTRING(@InputData,@Counter,1);
			SET @OneValue = ASCII(@TempChar);
			IF @OneValue >= 48 AND @OneValue <=57 
				SET @FilteredData = @FilteredData + @TempChar;
	        SET @Counter = @Counter+1;
	  END

      SET @FilteredLength = LEN(@FilteredData);	
	  IF @FilteredLength <= 0 
	  BEGIN
	      SET @ReturnData = ''; 
		  RETURN(@ReturnData);
	  END
	  
   SET @FilteredLength = LEN(@FilteredData);	
   IF @FilteredLength > 11 
    SET @FilteredData = SUBSTRING(@FilteredData,1,11);

   SET @FilteredLength = LEN(@FilteredData);	
   IF @FilteredLength < 11
   BEGIN
	SET @Lendiff=11-@FilteredLength;
	SET @Counter = 1;
	While @Counter<=@Lendiff
	BEGIN
	    SET @FilteredData = '0' + @FilteredData;	
		SET @Counter = @Counter+1;
	END
   END


   SET @CD=dbo.generateUPCACheckDigit(@FilteredData);
   SET @FilteredData=@FilteredData+@CD;

   SET @X = 1;
   While @X<=6
   BEGIN
		SET @Transformdataleft=@Transformdataleft + SUBSTRING(@FilteredData,@X,1);
	    SET @X = @X+1;   
   END

   SET @X = 7;
   While @X<=12
   BEGIN
	SET @Transformchar=SUBSTRING(@FilteredData,@X,1);
	SET @Transformvalue=ASCII(@Transformchar)+49; 
	SET @Transformdataright=@Transformdataright+dbo.UNICHR(@Transformvalue);
	SET @X = @X+1;   
   END 	   
   
   IF @Hr=1
		SET @ResultData=   dbo.UPCAUNICHR(ASCII(SUBSTRING(@Transformdataleft,1,1))-15) + '[' + dbo.UPCAUNICHR(ASCII(SUBSTRING(@Transformdataleft,1,1))+110) + SUBSTRING(@Transformdataleft,2,5) + '-' + SUBSTRING(@Transformdataright,1,5) + dbo.UPCAUNICHR(ASCII(SUBSTRING(@Transformdataright,6,1))-49+159) + ']' + dbo.UPCAUNICHR(ASCII(SUBSTRING(@Transformdataright,6,1))-49-15);
   ELSE
		SET @ResultData='[' + @Transformdataleft + '-' + @Transformdataright + ']';

   RETURN @ResultData;

  END
GO