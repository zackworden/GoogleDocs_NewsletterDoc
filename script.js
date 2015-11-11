// const
	var newsletterDocDetails = 	{
									name : 'Prototype Doc',
									year : 2016,
								};
// vars
	var thisSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	var ssActions = new SpreadsheetActions();
// enums
	var reportPeriods = {
							monthly	 : 'monthly',
							yearly	 : 'yearly',
						};
// classes
	function SpreadsheetReports()
	{
		this.AdvertiserReport = function()
		{
			
		}
		this.AllAdvertiserReport = function()
		{
			
		}
	}
	function SpreadsheetActions()
	{
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
	}
	function AdvertiserResultSet( theName,theDate, theSends, theOpens, theClicks, theSendType, adPosition )
	{
		this.advertiserName;
		this.allResults = [];
		
		this.Insert = function( theDate, theSends, theOpens, theClicks, theSendType, adPosition )
		{
			var tempResultItem = new ResultItem_(theDate, theSends, theOpens, theClicks, theSendType, adPosition);
			this.allResultItems.push(tempResultItem);
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
		var theSheet = thisSpreadsheet.getSheetByName('testSheet');
		
		ssActions.BuildWeeklyNewsletterDates(daysOfWeekArray, adPositionArray, theSheet, newsletterDocDetails.year);
		Browser.msgBox(newsletterDocDetails.name);
	}