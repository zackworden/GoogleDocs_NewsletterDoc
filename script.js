// const
	var newsletterDocDetails = 	{
									name : 'Prototype Doc',
									year : 2016,
									reportSheet : 'Reports',
								};
	var newsletterDocStructure	= 	{
										sendColumn : 1,
										openColumn : 2,
										dateColumn : 4,
										newsletterType : 5,
										adPositionColumn : 6,
										advertiserColumn : 7,
										clicksColumn : 9,
									};
// vars
	var thisSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	var ssActions = new SpreadsheetActions();
	var ssReports = new SpreadsheetReports();
	var ssReports2 = new Reports();
	var reportResults
// enums
	var reportPeriods = {
							monthly	 : 'monthly',
							yearly	 : 'yearly',
						};
// classes
	function ResultCollection()
	{
		this.advertiserResultSets = [];
		
		this.Get_advertiserResultSetIndex = function( advertiserName )
		{
			var resultSetIndex = -1;
			var counter = 0;
			var numOfResultSets = this.advertiserResultSets.length;
			
			for ( counter = 0; counter < numOfResultSets; counter ++ )
			{
				if ( this.advertiserResultSets[counter].advertiserName.toUpperCase() == advertiserName.toUpperCase() )
				{
					resultSetIndex = counter;
					break;
				}
			}
			
			return resultSetIndex;
		}
		this.AddAdvertiserResultSet = function( advertiserResultSet )
		{
			var theResultSetIndex = this.Get_advertiserResultSetIndex( advertiserResultSet.advertiserName );
			
			if ( theResultSetIndex == -1 )
			{
				this.advertiserResultSets.push( advertiserResultSet );
			}
			else
			{
				this.advertiserResultSets[theResultSetIndex].InsertResultSet( advertiserResultSet );
			}
		}
	}
	function Reports()
	{
		this.AdvertiserReportFullYear = function( theSheet, advertiserName )
		{
			var allResults = new ResultCollection();
			
			// consts
			var indexToColAdjustment = 1;
			
			// vars
			var counter = 0;
			var numOfRows = theSheet.getLastRow();
			var numOfCols = theSheet.getLastColumn();
			var theRange = theSheet.getRange(3,1,numOfRows,numOfCols);
			var theResults = theRange.getValues();
			var numOfResults = theResults.length;
			var thisAdvertiserResultSet;
			
			for ( counter = 0; counter < numOfResults; counter ++ )
			{
				if ( theResults[counter][(newsletterDocStructure.advertiserColumn - indexToColAdjustment)].toUpperCase() == advertiserName.toUpperCase() )
				{
					thisAdvertiserResultSet = new AdvertiserResultSet( theResults[counter][newsletterDocStructure.advertiserColumn - indexToColAdjustment] );
					thisAdvertiserResultSet.Insert(
						theResults[counter][newsletterDocStructure.dateColumn - indexToColAdjustment],
						theResults[counter][newsletterDocStructure.sendColumn - indexToColAdjustment], 
						theResults[counter][newsletterDocStructure.openColumn - indexToColAdjustment],
						theResults[counter][newsletterDocStructure.clicksColumn - indexToColAdjustment],
						theResults[counter][newsletterDocStructure.newsletterType - indexToColAdjustment],
						theResults[counter][newsletterDocStructure.adPositionColumn - indexToColAdjustment] 
						);
					
					this.allResults.AddAdvertiserResultSet( thisAdvertiserResultSet );
				}
			}
			
			this.BuildReport = function( reportSheet )
			{
				Logger.log( allResults[0] );
				
				
				
				var theRange = reportSheet.getRange( 1,1,reportSheet.getMaxRows(),reportSheet.getMaxColumns() );
				var advertiserCounter = 0;
				var numOfAdvertiserResults = allResults.length;
				var advertiserItemCounter = 0;
				var numOfAdvertiserItems = 0;
				var rangeAdjustment = 0;
				
				for ( advertiserCounter = 0; advertiserCounter < numOfAdvertiserResults; advertiserCounter ++ )
				{
					Logger.log( allResults[advertiserCounter].advertiserName );
					
					// create header
					theRange.getCell(1, 1).setValue('Advertiser');
					theRange.getCell(1, 2).setValue(allResults[advertiserCounter].advertiserName);
					theRange.getCell(2, 1).setValue('Date');
					theRange.getCell(2, 2).setValue('Publication');
					theRange.getCell(2, 3).setValue('Newsletter Type');
					theRange.getCell(2, 4).setValue('Sends');
					theRange.getCell(2, 5).setValue('Opens');
					theRange.getCell(2, 6).setValue('Clicks');
					theRange.getCell(2, 7).setValue('CTR');
					
					// for each, list advertiser itesm
					numOfAdvertiserItems = allResults[advertiserCounter].allResults.length;
					
					for ( advertiserItemCounter = 0; advertiserItemCounter < numOfAdvertiserItems; advertiserItemCounter ++ )
					{
						theRange.getCell(3 + advertiserItemCounter, 1).setValue( allResults[advertiserCounter].allResults[advertiserItemCounter].date );
						theRange.getCell(3 + advertiserItemCounter, 2).setValue( newsletterDocDetails.name );
						theRange.getCell(3 + advertiserItemCounter, 3).setValue( allResults[advertiserCounter].allResults[advertiserItemCounter].sendType );
						theRange.getCell(3 + advertiserItemCounter, 4).setValue( allResults[advertiserCounter].allResults[advertiserItemCounter].sends );
						theRange.getCell(3 + advertiserItemCounter, 5).setValue( allResults[advertiserCounter].allResults[advertiserItemCounter].opens );
						theRange.getCell(3 + advertiserItemCounter, 6).setValue( allResults[advertiserCounter].allResults[advertiserItemCounter].clicks );
						theRange.getCell(3 + advertiserItemCounter, 7).setFormula('=F' + ( advertiserItemCounter + 3) + '/E' + ( counter + 3) + '' );
						theRange.getCell(3 + advertiserItemCounter, 7).setNumberFormat('0.00');
						theRange.getCell(3 + advertiserItemCounter, 8).setValue( allResults[advertiserCounter].allResults[advertiserItemCounter].adPosition );
					}
					Logger.log('report built!');
				}
			}
		}
		this.AllAdvertiserReportFullYear = function()
		{
			this.allResults = [];
			
			this.BuildReport = function()
			{
				
			}
		}
		this.NewsletterReportFullYear = function()
		{
			this.allResults = [];
			
			this.BuildReport = function()
			{
				
			}
		}
	}
	function SpreadsheetReports()
	{
		this.allResultSets = [];
		
		// result gathering
		this.AdvertiserReport = function( theSheet, advertiserName )
		{
			// consts
			var indexToColAdjustment = 1;
			
			// vars
			var counter = 0;
			var numOfRows = theSheet.getLastRow();
			var numOfCols = theSheet.getLastColumn();
			var theRange = theSheet.getRange(3,1,numOfRows,numOfCols);
			var theResults = theRange.getValues();
			var numOfResults = theResults.length;
			var thisAdvertiserResultSet;
			
			for ( counter = 0; counter < numOfResults; counter ++ )
			{
				if ( theResults[counter][(newsletterDocStructure.advertiserColumn - indexToColAdjustment)].toUpperCase() == advertiserName.toUpperCase() )
				{
					thisAdvertiserResultSet = new AdvertiserResultSet( theResults[counter][newsletterDocStructure.advertiserColumn - indexToColAdjustment] );
					thisAdvertiserResultSet.Insert(
						theResults[counter][newsletterDocStructure.dateColumn - indexToColAdjustment],
						theResults[counter][newsletterDocStructure.sendColumn - indexToColAdjustment], 
						theResults[counter][newsletterDocStructure.openColumn - indexToColAdjustment],
						theResults[counter][newsletterDocStructure.clicksColumn - indexToColAdjustment],
						theResults[counter][newsletterDocStructure.newsletterType - indexToColAdjustment],
						theResults[counter][newsletterDocStructure.adPositionColumn - indexToColAdjustment] 
						);
					
					this.AddAdvertiserResultSet( thisAdvertiserResultSet );
				}
			}
		}
		this.AllAdvertiserReport = function( theSheet )
		{
			// consts
			var indexToColAdjustment = 1;
			
			// vars
			var counter = 0;
			var numOfRows = theSheet.getLastRow();
			var numOfCols = theSheet.getLastColumn();
			var theRange = theSheet.getRange(3,1,numOfRows,numOfCols);
			var theResults = theRange.getValues();
			var numOfResults = theResults.length;
			var thisAdvertiserResultSet;
			
			for ( counter = 0; counter < numOfResults; counter ++ )
			{
				thisAdvertiserResultSet = new AdvertiserResultSet( theResults[counter][newsletterDocStructure.advertiserColumn - indexToColAdjustment] );
				thisAdvertiserResultSet.Insert(
					theResults[counter][newsletterDocStructure.dateColumn - indexToColAdjustment],
					theResults[counter][newsletterDocStructure.sendColumn - indexToColAdjustment], 
					theResults[counter][newsletterDocStructure.openColumn - indexToColAdjustment],
					theResults[counter][newsletterDocStructure.clicksColumn - indexToColAdjustment],
					theResults[counter][newsletterDocStructure.newsletterType - indexToColAdjustment],
					theResults[counter][newsletterDocStructure.adPositionColumn - indexToColAdjustment] 
					);
				
				this.AddAdvertiserResultSet( thisAdvertiserResultSet );
			}
		}
		this.AddAdvertiserResultSet = function( theAdvertiserResultSet )
		{
			var theResultSetIndex = this.GetDoesAdvertiserResultSetExist( theAdvertiserResultSet.advertiserName );
			
			if ( theResultSetIndex == -1 )
			{
				this.allResultSets.push( theAdvertiserResultSet );
			}
			else
			{
				this.allResultSets[theResultSetIndex].InsertResultSet( theAdvertiserResultSet );
			}
		}
		this.GetDoesAdvertiserResultSetExist = function( theAdvertiserName )
		{
			var resultSetIndex = -1;
			var counter = 0;
			var numOfResultSets = this.allResultSets.length;
			
			for ( counter = 0; counter < numOfResultSets; counter ++ )
			{
				if ( this.allResultSets[counter].advertiserName.toUpperCase() == theAdvertiserName.toUpperCase() )
				{
					resultSetIndex = counter;
					break;
				}
			}
			
			return resultSetIndex;
		}
	
	}
	function SpreadsheetActions()
	{
		this.BuildReportSheet = function( theResultSetCollection )
		{
			var rowOffset = 0;
			var reportSheet = thisSpreadsheet.getSheetByName( newsletterDocDetails.reportSheet );
			var numOfRows = reportSheet.getMaxRows();
			var numOfCols = reportSheet.getMaxColumns();
			var reportRange = reportSheet.getRange(1,1,numOfRows, numOfCols);
			this.ClearSheet( reportSheet );
			
			var counter = 0;
			var numOfResultSets = theResultSetCollection.length;
			
			for ( counter = 0; counter < numOfResultSets; counter ++ )
			{
				this.BuildAdvertiserResultSet( reportRange, counter + rowOffset, theResultSetCollection[counter] );
				rowOffset += (theResultSetCollection[counter].allResults.length + 3);
			}
		}
		this.BuildWeeklyNewsletterDates = function( booleanArrayOfDaysOfWeek, arrayOfAdPositions, theSheet, currentYear )
		{
			// consts
			var theDateCol = 4;			// which column number maps to the 'date' column?
			var numOfDaysInYear = 365;	// hos many possible newsletters could go in a year, under the probably bad assumption that 1 / day is the most we will do.
			
			// vars
			var adPositionCounter = 0;
			var numOfAdPositions = arrayOfAdPositions.length;
			var counter = 0;
			var rowCounter = 1;
			var theSheetRange = theSheet.getRange(theDateCol,3,1500,2);
			var theDate = new Date();
			
			for ( counter = 0; counter < numOfDaysInYear; counter ++ )
			{
				theDate = new Date( currentYear, 0, counter );
				
				if ( booleanArrayOfDaysOfWeek[ theDate.getDay() ] == true )
				{
					for ( adPositionCounter = 0; adPositionCounter < numOfAdPositions; adPositionCounter ++ )
					{
						theSheetRange.getCell( rowCounter + adPositionCounter, 1 ).setValue( theDate );
						theSheetRange.getCell( rowCounter + adPositionCounter, 2 ).setValue( arrayOfAdPositions[adPositionCounter] );
					}
					
					rowCounter += (numOfAdPositions + 1);
				}
			}
		}
		this.ClearSheet = function( theSheet )
		{
			theSheet.clear();
		}
		
		this.BuildAdvertiserResultSet = function( theRange, rangeAdjustment, theResultSet )
		{
			var counter = 0;
			var numOfResultItems = theResultSet.allResults.length;
			var resultSetRange = theRange.offset( rangeAdjustment,0, (rangeAdjustment + numOfResultItems + 2) );
			var resultItemRange;
			
			// advertiser result set header
			resultSetRange.getCell(1, 1).setValue('Advertiser');
			resultSetRange.getCell(1, 2).setValue(theResultSet.advertiserName);
			resultSetRange.getCell(2, 1).setValue('Date');
			resultSetRange.getCell(2, 2).setValue('Publication');
			resultSetRange.getCell(2, 3).setValue('Newsletter Type');
			resultSetRange.getCell(2, 4).setValue('Sends');
			resultSetRange.getCell(2, 5).setValue('Opens');
			resultSetRange.getCell(2, 6).setValue('Clicks');
			resultSetRange.getCell(2, 7).setValue('CTR');
			
			//  advertiser result body
			for ( counter = 0; counter < numOfResultItems; counter ++ )
			{
				resultSetRange.getCell(3 + counter, 1).setValue(theResultSet.allResults[counter].date);
				resultSetRange.getCell(3 + counter, 2).setValue( newsletterDocDetails.name );
				resultSetRange.getCell(3 + counter, 3).setValue(theResultSet.allResults[counter].sendType);
				resultSetRange.getCell(3 + counter, 4).setValue(theResultSet.allResults[counter].sends);
				resultSetRange.getCell(3 + counter, 5).setValue(theResultSet.allResults[counter].opens);
				resultSetRange.getCell(3 + counter, 6).setValue(theResultSet.allResults[counter].clicks);
				resultSetRange.getCell(3 + counter, 7).setFormula('=F' + (rangeAdjustment + counter + 3) + '/E' + (rangeAdjustment + counter + 3) + '');
				resultSetRange.getCell(3 + counter, 7).setNumberFormat('0.00');
				resultSetRange.getCell(3 + counter, 8).setValue(theResultSet.allResults[counter].adPosition);
				theResultSet.allResults[counter].date
			}
		}
		this.BuildAdvertiserResultSetItem = function( theRange, theResultSetRow )
		{
			
		}
	}
	function AdvertiserResultSet( theName )
	{
		this.advertiserName = theName;
		this.allResults = [];
		
		this.InsertResultSet = function( theAdvertiserResultSet )
		{
			this.Insert( 	
				theAdvertiserResultSet.allResults[0].date, 
				theAdvertiserResultSet.allResults[0].sends, 
				theAdvertiserResultSet.allResults[0].opens, 
				theAdvertiserResultSet.allResults[0].clicks, 
				theAdvertiserResultSet.allResults[0].sendType, 
				theAdvertiserResultSet.allResults[0].adPosition 
				);
		}
		this.Insert = function( theDate, theSends, theOpens, theClicks, theSendType, adPosition )
		{
			var tempResultItem = new AdvertiserResultSetItem(theDate, theSends, theOpens, theClicks, theSendType, adPosition);
			this.allResults.push(tempResultItem);
		}
	}
	function AdvertiserResultSetItem( theDate, theSends, theOpens, theClicks, theSendType, adPosition )
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
		var daysOfWeekArray = [false, false, false, true, true, false, false];
		var adPositionArray = ['Top','Middle'];
		var theSheet = thisSpreadsheet.getSheetByName('Sheet1');
		
		//ssActions.BuildWeeklyNewsletterDates(daysOfWeekArray, adPositionArray, theSheet, newsletterDocDetails.year);
		/*
		ssReports.AllAdvertiserReport(theSheet);
		
		ssActions.BuildReportSheet( ssReports.allResultSets );
		*/
		ssReports2.AdvertiserReportFullYear(theSheet, 'Zack');
		ssReports2.BuildReport( thisSpreadsheet.getSheetByName(newsletterDocDetails.reportSheet) );
	}