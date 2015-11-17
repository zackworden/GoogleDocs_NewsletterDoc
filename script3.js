// test
// const
	var newsletterMeta = 	{
								year : 2016,
								publication : 'Prototype',
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
	var sheet_test =  thisSpreadsheet.getSheetByName('testSheet');
	var sheet_test2 =  thisSpreadsheet.getSheetByName('Product Showcase');
	var sheet_report =  thisSpreadsheet.getSheetByName('Reports');
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
				this.ProcessRow( theValues[counter], advertiserName, dateFilter, theSheet.getName() );
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
				this.ProcessRow( theValues[counter], 'ALL', dateFilter, theSheet.getName() );
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
				theRange.getCell( rowNumber + 2 + counter, 2).setValue( thisPlacement.publication );
				theRange.getCell( rowNumber + 2 + counter, 3).setValue( thisPlacement.sendType );
				theRange.getCell( rowNumber + 2 + counter, 4).setValue( thisPlacement.sends );
				theRange.getCell( rowNumber + 2 + counter, 5).setValue( thisPlacement.opens );
				theRange.getCell( rowNumber + 2 + counter, 6).setValue( thisPlacement.clicks );
				theRange.getCell( rowNumber + 2 + counter, 7).setFormulaR1C1( '=R[0]C[-1]/R[0]C[-2]' );
				theRange.getCell( rowNumber + 2 + counter, 7).setNumberFormat( '0.00' );
				theRange.getCell( rowNumber + 2 + counter, 8).setValue( thisPlacement.adPosition );
			}
		}
		this.ProcessRow = function( theValueRow, theAdvertiserName, theDateFilter, sheetName )	// treat as private
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
																			theValueRow[newsletterDocStructure.newsletterType - 1],
																			theValueRow[newsletterDocStructure.adPositionColumn - 1],
																			newsletterMeta.publication
																		);
				// check to see if a newsletter type has been set. if so, use that. if not, use the sheet name.
					if ( theValueRow[(newsletterDocStructure.newsletterType - 1)].length > 2 )
					{
						thisPlacementResult.sendType = theValueRow[(newsletterDocStructure.newsletterType - 1)];
					}
					else
					{
						thisPlacementResult.sendType = sheetName;
					}
				
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
				newPlacement = new Advertiser_Placement_Result( theAdvertiserResult.allPlacements[0].date, theAdvertiserResult.allPlacements[0].sends, theAdvertiserResult.allPlacements[0].opens, theAdvertiserResult.allPlacements[0].clicks, theAdvertiserResult.allPlacements[0].sendType, theAdvertiserResult.allPlacements[0].adPosition, theAdvertiserResult.allPlacements[0].publication );
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
	function Advertiser_Placement_Result( theDate, theSends, theOpens, theClicks, theSendType, theAdPosition, publicationName )
	{
		this.date = theDate;
		this.sends = theSends;
		this.opens = theOpens;
		this.clicks = theClicks;
		this.sendType = theSendType;
		this.adPosition = theAdPosition;
		this.publication = publicationName;
	}
// functions
	function onOpen()
	{
		var ui = SpreadsheetApp.getUi();
		ui.createMenu('Reporting')
			.addSubMenu( 
						ui.createMenu('By Advertiser') 
							.addItem( 'All 2016', 'Menu_AdReport_2016_' )
							.addItem( 'January', 'Menu_AdReport_Jan_' )
							.addItem( 'February', 'Menu_AdReport_Feb_' )
							.addItem( 'March', 'Menu_AdReport_Mar_' )
							.addItem( 'April', 'Menu_AdReport_Apr_' )
							.addItem( 'May', 'Menu_AdReport_May_' )
							.addItem( 'June', 'Menu_AdReport_Jun_' )
							.addItem( 'July', 'Menu_AdReport_Jul_' )
							.addItem( 'August', 'Menu_AdReport_Aug_' )
							.addItem( 'September', 'Menu_AdReport_Sep_' )
							.addItem( 'October', 'Menu_AdReport_Oct_' )
							.addItem( 'November', 'Menu_AdReport_Nov_' )
							.addItem( 'December', 'Menu_AdReport_Dec_' )
							)
			.addSeparator()
			.addSubMenu( 
						ui.createMenu('All Advertisers') 
							.addItem( 'All 2016', 'Menu_AllAdReport_2016_' )
							.addItem( 'January', 'Menu_AllAdReport_Jan_' )
							.addItem( 'February', 'Menu_AllAdReport_Feb_' )
							.addItem( 'March', 'Menu_AllAdReport_Mar_' )
							.addItem( 'April', 'Menu_AllAdReport_Apr_' )
							.addItem( 'May', 'Menu_AllAdReport_May_' )
							.addItem( 'June', 'Menu_AllAdReport_Jun_' )
							.addItem( 'July', 'Menu_AllAdReport_Jul_' )
							.addItem( 'August', 'Menu_AllAdReport_Aug_' )
							.addItem( 'September', 'Menu_AllAdReport_Sep_' )
							.addItem( 'October', 'Menu_AllAdReport_Oct_' )
							.addItem( 'November', 'Menu_AllAdReport_Nov_' )
							.addItem( 'December', 'Menu_AllAdReport_Dec_' )
							)
			.addSeparator()
			.addToUi();
	}
	function Menu_AdReport_2016_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 0, 1),
								endDate : new Date(newsletterMeta.year, 11, 0),
							};
		Create_AdvertiserReport_( dateFilter );
	}
	function Menu_AdReport_Jan_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 0, 1),
								endDate : new Date(newsletterMeta.year, 1, 0),
							};
		Create_AdvertiserReport_( dateFilter );
	}
	function Menu_AdReport_Feb_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 1, 1),
								endDate : new Date(newsletterMeta.year, 2, 0),
							};
		Create_AdvertiserReport_( dateFilter );
	}
	function Menu_AdReport_Mar_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 2, 1),
								endDate : new Date(newsletterMeta.year, 3, 0),
							};
		Create_AdvertiserReport_( dateFilter );
	}
	function Menu_AdReport_Apr_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 3, 1),
								endDate : new Date(newsletterMeta.year, 4, 0),
							};
		Create_AdvertiserReport_( dateFilter );
	}
	function Menu_AdReport_May_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 4, 1),
								endDate : new Date(newsletterMeta.year, 5, 0),
							};
		Create_AdvertiserReport_( dateFilter );
	}
	function Menu_AdReport_Jun_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 5, 1),
								endDate : new Date(newsletterMeta.year, 6, 0),
							};
		Create_AdvertiserReport_( dateFilter );
	}
	function Menu_AdReport_Jul_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 6, 1),
								endDate : new Date(newsletterMeta.year, 7, 0),
							};
		Create_AdvertiserReport_( dateFilter );
	}
	function Menu_AdReport_Aug_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 7, 1),
								endDate : new Date(newsletterMeta.year, 8, 0),
							};
		Create_AdvertiserReport_( dateFilter );
	}
	function Menu_AdReport_Sep_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 8, 1),
								endDate : new Date(newsletterMeta.year, 9, 0),
							};
		Create_AdvertiserReport_( dateFilter );
	}
	function Menu_AdReport_Oct_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 9, 1),
								endDate : new Date(newsletterMeta.year, 10, 0),
							};
		Create_AdvertiserReport_( dateFilter );
	}
	function Menu_AdReport_Nov_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 10, 1),
								endDate : new Date(newsletterMeta.year, 11, 0),
							};
		Create_AdvertiserReport_( dateFilter );
	}
	function Menu_AdReport_Dec_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 11, 1),
								endDate : new Date(newsletterMeta.year, 12, 0),
							};
		Create_AdvertiserReport_( dateFilter );
	}
	function Menu_AllAdReport_2016_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 0, 1),
								endDate : new Date(newsletterMeta.year, 11, 0),
							};
		Create_AllAdvertiserReport( dateFilter );
	}
	function Menu_AllAdReport_Jan_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 0, 1),
								endDate : new Date(newsletterMeta.year, 1, 0),
							};
		Create_AllAdvertiserReport( dateFilter );
	}
	function Menu_AllAdReport_Feb_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 1, 1),
								endDate : new Date(newsletterMeta.year, 2, 0),
							};
		Create_AllAdvertiserReport( dateFilter );
	}
	function Menu_AllAdReport_Mar_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 2, 1),
								endDate : new Date(newsletterMeta.year, 3, 0),
							};
		Create_AllAdvertiserReport( dateFilter );
	}
	function Menu_AllAdReport_Apr_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 3, 1),
								endDate : new Date(newsletterMeta.year, 4, 0),
							};
		Create_AllAdvertiserReport( dateFilter );
	}
	function Menu_AllAdReport_May_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 4, 1),
								endDate : new Date(newsletterMeta.year, 5, 0),
							};
		Create_AllAdvertiserReport( dateFilter );
	}
	function Menu_AllAdReport_Jun_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 5, 1),
								endDate : new Date(newsletterMeta.year, 6, 0),
							};
		Create_AllAdvertiserReport( dateFilter );
	}
	function Menu_AllAdReport_Jul_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 6, 1),
								endDate : new Date(newsletterMeta.year, 7, 0),
							};
		Create_AllAdvertiserReport( dateFilter );
	}
	function Menu_AllAdReport_Aug_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 7, 1),
								endDate : new Date(newsletterMeta.year, 8, 0),
							};
		Create_AllAdvertiserReport( dateFilter );
	}
	function Menu_AllAdReport_Sep_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 8, 1),
								endDate : new Date(newsletterMeta.year, 9, 0),
							};
		Create_AllAdvertiserReport( dateFilter );
	}
	function Menu_AllAdReport_Oct_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 9, 1),
								endDate : new Date(newsletterMeta.year, 10, 0),
							};
		Create_AllAdvertiserReport( dateFilter );
	}
	function Menu_AllAdReport_Nov_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 10, 1),
								endDate : new Date(newsletterMeta.year, 11, 0),
							};
		Create_AllAdvertiserReport( dateFilter );
	}
	function Menu_AllAdReport_Dec_()
	{
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 11, 1),
								endDate : new Date(newsletterMeta.year, 12, 0),
							};
		Create_AllAdvertiserReport( dateFilter );
	}
	function Create_AdvertiserReport_( theDateFilter )
	{
		var adReport = new AdvertiserReport();
		var theUI = SpreadsheetApp.getUi();
		var advertiserResponse = theUI.prompt('Which advertiser are you looking up?','',theUI.ButtonSet.OK_CANCEL);
		var advertiserName;
		
		if ( advertiserResponse.getSelectedButton() === theUI.Button.OK )
		{
			advertiserName = advertiserResponse.getResponseText();
			advertiserName.toUpperCase();
			
			adReport.GetAdvertiserResults( sheet_test, advertiserName, theDateFilter );
			adReport.GetAdvertiserResults( sheet_test2, advertiserName, theDateFilter );
			adReport.BuildReport( sheet_report );
		}
	}
	function Create_AllAdvertiserReport( theDateFilter )
	{
		var adReport = new AdvertiserReport();
		adReport.GetAdvertiserResults( sheet_test, 'ALL', theDateFilter );
		adReport.GetAdvertiserResults( sheet_test2, 'ALL', theDateFilter );
		adReport.BuildReport( sheet_report );
	}
	function Test()
	{
		var theSheet = thisSpreadsheet.getSheetByName('testSheet');
		var reportSheet = thisSpreadsheet.getSheetByName('Reports');
		var adReport = new AdvertiserReport();
		
		var dateFilter = 	{
								startDate : new Date(newsletterMeta.year, 0, 1),
								endDate : new Date(newsletterMeta.year, 1, 0),
							};
		adReport.GetAdvertiserResults(theSheet, 'zack', dateFilter);
		adReport.GetAdvertiserResults(theSheet, 'lutron', dateFilter);
		adReport.BuildReport(reportSheet);
	}