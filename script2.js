// consts
	var newsletterDocDetails = 	{
									name : 'Prototype Doc',
									year : 2016,
									reportSheet : 'Reports',
								};
	var newsletterDocStructure	= 	{
										sendColumn : 1,
										openColumn : 2,
										openRateColumn : 3,
										dateColumn : 4,
										newsletterType : 5,
										adPositionColumn : 6,
										advertiserColumn : 7,
										revenueColumn : 8,
										clicksColumn : 9,
									};
// vars
	var thisSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	var sheet_test = thisSpreadsheet.getSheetByName('Sheet1');
	var sheet_report = thisSpreadsheet.getSheetByName('Reports');
	var AdvertiserReport = new AdvertiserReports();
// enums
// classes
	function AdvertiserReports()
	{
		this.resultCollection = new ResultCollection();
		
		this.GetResults = function( theSheet, advertiserName )
		{
			var indexToColAdjustment = 1;			// JS arrays start at 0, while google cells start at 1
			var counter = 0;
			var firstRow = 3;
			var firstCol = 1;
			var numOfRows = theSheet.getLastRow();
			var numOfCols = theSheet.getLastColumn();
			var theRange = theSheet.getRange(firstRow,firstCol,numOfRows - (firstRow - 1), numOfCols - (firstCol - 1) );
			var theValues = theRange.getValues();
			var numOfValues = theValues.length;
			var thisResultSet;
			
			for ( counter = 0; counter < numOfValues; counter ++ )
			{
				if ( theValues[counter][newsletterDocStructure.advertiserColumn - indexToColAdjustment].toUpperCase() == advertiserName.toUpperCase() )
				{
					thisResultSet = new ResultSet( theValues[counter][newsletterDocStructure.advertiserColumn - indexToColAdjustment].toUpperCase() );
					thisResultSet.AddResultItem	(
													theValues[counter][newsletterDocStructure.dateColumn - indexToColAdjustment],
													theValues[counter][newsletterDocStructure.sendColumn - indexToColAdjustment],
													theValues[counter][newsletterDocStructure.openColumn - indexToColAdjustment],
													theValues[counter][newsletterDocStructure.clicksColumn - indexToColAdjustment],
													theValues[counter][newsletterDocStructure.newsletterType - indexToColAdjustment],
													theValues[counter][newsletterDocStructure.adPositionColumn - indexToColAdjustment]
												);
					this.resultCollection.AddResultSet( thisResultSet );
				}
			}
			//Logger.log( this.resultCollection.allResultSets[0].allResultItems );
		}
		this.BuildResults = function( reportSheet )
		{
			// test
			var thisResultItem;
			
			// real
			var firstCol = 1;
			var numOfCols = 8;
			var firstRow = 1;
			var numOfRows = reportSheet.getMaxRows();
			var theRange = reportSheet.getRange(firstRow,firstCol,numOfRows, numOfCols );
			
			// resultSet header
			
			var advertiserCounter = 0;
			var numOfResultSets = this.resultCollection.allResultSets.length;
			
			for ( advertiserCounter = 0; advertiserCounter < numOfResultSets; advertiserCounter ++ )
			{
				var thisResultSet =  this.resultCollection.allResultSets[advertiserCounter];
				
				
				theRange.getCell(1, 1).setValue('Advertiser');
				theRange.getCell(1, 2).setValue(thisResultSet.advertiserName);
				theRange.getCell(2, 1).setValue('Date');
				theRange.getCell(2, 2).setValue('Publication');
				theRange.getCell(2, 3).setValue('Newsletter Type');
				theRange.getCell(2, 4).setValue('Sends');
				theRange.getCell(2, 5).setValue('Opens');
				theRange.getCell(2, 6).setValue('Clicks');
				theRange.getCell(2, 7).setValue('CTR');
				
				var resultItemCounter = 0;
				var numOfResultItems = thisResultSet.allResultItems.length;
				
				for ( resultItemCounter = 0; resultItemCounter < numOfResultItems; resultItemCounter ++ )
				{
					thisResultItem = thisResultSet.allResultItems[resultItemCounter];
					
					theRange.getCell(3 + resultItemCounter, 1).setValue( thisResultItem.date );
					theRange.getCell(3 + resultItemCounter, 2).setValue( newsletterDocDetails.name );
					theRange.getCell(3 + resultItemCounter, 3).setValue( thisResultItem.sendType );
					theRange.getCell(3 + resultItemCounter, 4).setValue( thisResultItem.sends );
					theRange.getCell(3 + resultItemCounter, 5).setValue( thisResultItem.opens );
					theRange.getCell(3 + resultItemCounter, 6).setValue( thisResultItem.clicks );
					theRange.getCell(3 + resultItemCounter, 7).setFormula('=F' + ( resultItemCounter + 3) + '/E' + ( resultItemCounter + 3) + '' );
					theRange.getCell(3 + resultItemCounter, 7).setNumberFormat('0.00');
					theRange.getCell(3 + resultItemCounter, 8).setValue( thisResultItem.adPosition );
				}
			}
		}
	}
	function ResultCollection()
	{
		this.allResultSets = [];
		
		this.AddResultSet = function( theResultSet )
		{
			var counter = 0;
			var numOfResultSets = this.allResultSets.length;
			var wasResultFound = false;
			
			for ( counter = 0; counter < numOfResultSets; counter ++ )
			{
				if ( this.allResultSets[counter].advertiserName.toUpperCase() == theResultSet.advertiserName.toUpperCase() )
				{
					wasResultFound = true;
					this.allResultSets[counter].AddResultSet( theResultSet );
					break;
				}
			}
			if ( wasResultFound === false )
			{
				this.allResultSets.push( theResultSet );
			}
			Logger.log( this.allResultSets[0].allResultItems );
		}
	}
	function ResultSet( theName )
	{
		this.allResultItems = [];
		this.advertiserName = theName;
		
		this.AddResultItem = function( theDate, theSends, theOpens, theClicks, theSendType, adPosition )
		{
			var tempResultItem = new ResultItem(theDate, theSends, theOpens, theClicks, theSendType, adPosition);
			this.allResultItems.push(tempResultItem);
		}
		this.AddResultSet = function( theResultSet )
		{
			this.AddResultItem( theResultSet.date, theResultSet.sends, theResultSet.opens, theResultSet.clicks, theResultSet.sendType, theResultSet.adPosition );
			Logger.log( theResultSet );
			Logger.log( '\r\n\r\n' );
			Logger.log( this.allResultItems );
		}
	}
	function ResultItem( theDate, theSends, theOpens, theClicks, theSendType, adPosition )
	{
		this.date = theDate;
		this.sends = theSends;
		this.opens = theOpens;
		this.clicks = theClicks;
		this.sendType = theSendType;
		this.adPosition = adPosition;
	}
// functions
	function Test()
	{
		AdvertiserReport.GetResults(sheet_test, 'Zack');
		AdvertiserReport.BuildResults( sheet_report );
	}