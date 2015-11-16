// test
// const
	var newsletterMeta = 	{
								year : 2016,
							};
	var thisSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
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
// classes
	function AdvertiserReport()
	{
		this.allResults = new ResultCollection();	// : ResultCollection
		
		this.GetAdvertiserResults = function( theSheet, advertiserName, dateFilter )
		{
			var startingRow = 3;
			var startingCol = 1;
			var numOfRows = theSheet.getLastRow();
			var numOfCols = theSheet.getLastColumn();
			var theRange = theSheet.getRange( startingRow, startingCol, numOfRows - (startingRow - 1), numOfCols - (startingCol - 1) );
			
			var theValues = theRange.getValues();
			var counter = 0;
			var numOfValues = theValues.length;
			
			for ( counter = 0; counter < numOfValues; counter ++ )
			{
				this.ProcessRow( theValues[counter], advertiserName, dateFilter );
			}
		}
		this.GetAllResults = function( theSheet, dateFilter )
		{
			var startingRow = 3;
			var startingCol = 1;
			var numOfRows = theSheet.getLastRow();
			var numOfCols = theSheet.getLastColumn();
			var theRange = theSheet.getRange( startingRow, startingCol, numOfRows - (startingRow - 1), numOfCols - (startingCol - 1) );
			
			var theValues = theRange.getValues();
			var counter = 0;
			var numOfValues = theValues.length;
			
			for ( counter = 0; counter < numOfValues; counter ++ )
			{
				this.ProcessRow( theValues[counter], 'ALL', dateFilter );
			}
		}
		this.BuildReport = function( theSheet )
		{
			theSheet.clear();
			var firstCol = 1;
			var firstRow = 1;
			var lastCol = theSheet.getMaxColumns();
			var lastRow = theSheet.getMaxRows();
			var theRange = theSheet.getRange( firstRow, firstCol, (lastRow - firstRow), (lastCol - firstCol) );
			var advertiserRange;
			
			var counter = 0;
			var rowCounter = firstRow;
			var numOfAdvertisers = this.allResults.allAdvertiserResults.length;
			var thisAdvertiser;
			
			for ( counter = 0; counter < numOfAdvertisers; counter ++ )
			{
				thisAdvertiser = this.allResults.allAdvertiserResults[counter];
				this.BuildAdvertiserSection( thisAdvertiser, theRange, rowCounter );
				rowCounter += thisAdvertiser.allPlacements.length + 3;
			}
		}
		this.BuildAdvertiserSection = function( theAdvertiser, theRange, rowNumber )	// treat as private
		{
			// advertiser header
			theRange.getCell( rowNumber, 1).setValue('Advertiser Name');
			theRange.getCell( rowNumber, 2).setValue(theAdvertiser.advertiserName);
			theRange.getCell( rowNumber + 1, 1).setValue('Date');
			theRange.getCell( rowNumber + 1, 2).setValue('Publication');
			theRange.getCell( rowNumber + 1, 3).setValue('Newsletter Type');
			theRange.getCell( rowNumber + 1, 4).setValue('Sends');
			theRange.getCell( rowNumber + 1, 5).setValue('Opens');
			theRange.getCell( rowNumber + 1, 6).setValue('Clicks');
			theRange.getCell( rowNumber + 1, 7).setValue('CTR');
			theRange.getCell( rowNumber + 1, 8).setValue('Ad Position');
			
			// advertiser results
			var counter = 0;
			var numOfAds = theAdvertiser.allPlacements.length;
			var thisPlacement;
			
			for ( counter = 0; counter < numOfAds; counter ++ )
			{
				thisPlacement = theAdvertiser.allPlacements[counter];

				theRange.getCell( rowNumber + 2 + counter, 1).setValue( thisPlacement.date );
				theRange.getCell( rowNumber + 2 + counter, 2).setValue( 'publication name' );
				theRange.getCell( rowNumber + 2 + counter, 3).setValue( 'newsletter type' );
				theRange.getCell( rowNumber + 2 + counter, 4).setValue( thisPlacement.sends );
				theRange.getCell( rowNumber + 2 + counter, 5).setValue( thisPlacement.opens );
				theRange.getCell( rowNumber + 2 + counter, 6).setValue( thisPlacement.clicks );
				theRange.getCell( rowNumber + 2 + counter, 7).setValue( 'ctr' );
				theRange.getCell( rowNumber + 2 + counter, 8).setValue( thisPlacement.adPosition );
			}
		}
		this.ProcessRow = function( theValueRow, theAdvertiserName, theDateFilter )	// treat as private
		{
			var thisAdvertiserResult;
			var thisPlacementResult;
			
			// is row valid?
			if ( this.Get_IsValidResult( theValueRow, theAdvertiserName, theDateFilter ) === true )
			{
				// add
				thisAdvertiserResult = new Advertiser_Result();
				thisAdvertiserResult.advertiserName = theValueRow[newsletterDocStructure.advertiserColumn - 1].toUpperCase();
				thisPlacementResult = new Advertiser_Placement_Result	(
																			theValueRow[newsletterDocStructure.dateColumn - 1],
																			theValueRow[newsletterDocStructure.sendColumn - 1],
																			theValueRow[newsletterDocStructure.openColumn - 1],
																			theValueRow[newsletterDocStructure.clicksColumn - 1],
																			theValueRow[newsletterDocStructure.dateColumn - 1],
																			theValueRow[newsletterDocStructure.adPositionColumn - 1]
																		);
				thisAdvertiserResult.Add_PlacementResult( thisPlacementResult );
				
				this.allResults.Add_AdvertiserResult( thisAdvertiserResult );
				
				return true;
			}
			else
			{
				return false;
			}
		}
		this.Get_IsValidResult = function( theValueRow, theAdvertiserName, theDateFilter ) // treat as private. this evaluates each column, keeping all criteria within this function
		{
			var isValid = true;
			
			if ( !theValueRow[(newsletterDocStructure.advertiserColumn - 1)] )
			{
				isValid = false;
			}
			if ( theValueRow[(newsletterDocStructure.advertiserColumn - 1)] == '' )
			{
				isValid = false;
			}
			if ( theValueRow[(newsletterDocStructure.advertiserColumn - 1)] == ' ' )
			{
				isValid = false;
			}
			if ( theAdvertiserName.toUpperCase() != 'ALL' )
			{
				if ( theValueRow[(newsletterDocStructure.advertiserColumn - 1)].toUpperCase() != theAdvertiserName.toUpperCase() )
				{
					isValid = false;
				}
			}
			if ( theValueRow[(newsletterDocStructure.dateColumn - 1)] < theDateFilter.startDate )
			{
				isValid = false;
			}
			if ( theValueRow[(newsletterDocStructure.dateColumn - 1)] > theDateFilter.endDate )
			{
				isValid = false;
			}
			
			return isValid;
		}
	}
	function ResultCollection()
	{
		this.allAdvertiserResults = [];	// : Advertiser_Result
		
		this.Get_AdvertiserResultByName = function( advertiserName )
		{
			var counter = 0;
			var numOf = this.allAdvertiserResults.length;
			
			for ( counter = 0; counter < numOf; counter ++ )
			{
				if ( this.allAdvertiserResults[counter].advertiserName == advertiserName )
				{
					return this.allAdvertiserResults[counter];
				}
			}
			
			return -1;
		}
		this.Add_AdvertiserResult = function( theAdvertiserResult ) // : Advertiser_Result
		{
			var advertiser = this.Get_AdvertiserResultByName( theAdvertiserResult.advertiserName );
			
			if ( advertiser === -1 )
			{
				//Logger.log('adding a new advertiser report');
				this.allAdvertiserResults.push( theAdvertiserResult );
			}
			else
			{
				//Logger.log('adding to existing advertiser report');
				advertiser.Add_PlacementResultFromAdResult( theAdvertiserResult );
			}
		}
	}
	function Advertiser_Result()
	{
		this.advertiserName = 'default advertiser';
		this.allPlacements = [];	// : Advertiser_Placement_Result
		
		this.Add_PlacementResultFromAdResult = function( theAdvertiserResult )
		{
			var newPlacement;
			
			if ( theAdvertiserResult.allPlacements.length > 0 )
			{
				newPlacement = new Advertiser_Placement_Result( theAdvertiserResult.allPlacements[0].date, theAdvertiserResult.allPlacements[0].sends, theAdvertiserResult.allPlacements[0].opens, theAdvertiserResult.allPlacements[0].clicks, theAdvertiserResult.allPlacements[0].sendType, theAdvertiserResult.allPlacements[0].adPosition );
				this.Add_PlacementResult(newPlacement);
			}
			else
			{
				Logger.log('Error in Add_PlacementResultFromAdResult');
			}
		}
		this.Add_PlacementResult = function( thePlacement )  // : Advertiser_Placement_Result
		{
			this.allPlacements.push( thePlacement );
		}
	}
	function Advertiser_Placement_Result( theDate, theSends, theOpens, theClicks, theSendType, theAdPosition )
	{
		this.date = theDate;
		this.sends = theSends;
		this.opens = theOpens;
		this.clicks = theClicks;
		this.sendType = theSendType;
		this.adPosition = theAdPosition;
	}
// functions
	function Test()
	{
		var theSheet = thisSpreadsheet.getSheetByName('testSheet');
		var reportSheet = thisSpreadsheet.getSheetByName('Reports');
		var adReport = new AdvertiserReport();
		//adReport.GetAdvertiserResults(theSheet, 'zack');
		
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 0, 1),
								endDate : new Date(newsletterMeta.year, 1, 0),
							};
		//adReport.GetAllResults(theSheet, dateFilter);
		adReport.GetAdvertiserResults(theSheet, 'zack', dateFilter);
		adReport.BuildReport(reportSheet);
	}