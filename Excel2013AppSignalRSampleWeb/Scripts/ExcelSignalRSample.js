// This function is run when the app is ready to start interacting with the host application
// It ensures the DOM is ready before adding click handlers to buttons
Office.initialize = function (reason) {
    $(document).ready(function () {

        var stockProperties = ["Symbol", "Price", "DayHigh", "DayLow", "DayOpen", "Change", "LastChange", "PercentChange"];
        var tableBinding;
        var stockRowMap = {};

        $("#setBinding").click(function () {
            
            Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Table,
                // It would be nice if restrictions on the Binding parameters could be optionally specified here,
                // such as the 'shape' of a TableBinding/MatrixBinding (i.e. min/max number of rows/columns), or the
                // expected headers of a table's columns. As it is, we will check that here...
                { id: 'StockTable', promptText: 'Select a table to bind to.' },
                function (asyncResult) {
                    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                        $("#Messages").prepend("<p>An error occurred attempting to set the TableBinding: " +
                            asyncResult.error.message + "</p>");
                        return;
                    }

                    tableBinding = asyncResult.value;
                    
                    // In Excel 2013, it seems you can't create a table without headers, so a call to binding.hasHeaders
                    // isn't necessary, we can just get the headers async.
                    tableBinding.getDataAsync({ rowCount: 0 }, function (asyncResult) {
                        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                            $("#Messages").prepend("<p>An error occurred attempting to get the table headers: " +
                                asyncResult.error.message + "</p>");
                            return;
                        }

                        var headers = asyncResult.value.headers[0];
                        
                        // Determine if there are any unrecognised headers...
                        var unrecognisedHeaders = headers.filter(function (element) {
                            return stockProperties.indexOf(element) < 0;
                        });
                        
                        // and if so... create an error message and release the binding.
                        if (unrecognisedHeaders.length != 0) {
                            $("#Messages").append("<p>There are unrecognised column headers; please remove and create the binding again: " +
                                unrecognisedHeaders.join(",") + "</p>");

                            // Releasing the binding could fail... but not checking that at the moment.
                            Office.context.document.bindings.releaseByIdAsync('StockTable');
                            return;
                        }

                        // Otherwise, override the default mapStockToArray function with one that pushes the 
                        // matching property values of the received stock object into an array, in the same
                        // order as the TableBinding's column headings.
                        mapStockToArray = function(stock) {
                            var stockArray = [];
                            $.each(headers, function(idx, val) {
                                stockArray.push(stock[val]);
                            });
                            return stockArray;
                        };
                        initializeTable(tableBinding);
                    });
                });
        });

        $("#clearBinding").click(function() {

        });
        
        function initializeTable(binding) {
            ticker.server.getAllStocks().done(function(stocks) {
                $(stocks).each(function(idx, val) {
                    stockRowMap[val.Symbol] = idx;
                });
                stocks = stocks.map(mapStockToArray);

                binding.addRowsAsync(stocks, function(asyncResult) {
                    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {

                    }
                });
            });
        }
        
        var ticker = $.connection.stockTicker;

        function mapStockToArray() {
            $('#Messages').append("<p>mapStockToArray function has not yet been defined.</p>");
        }

        $.extend(ticker.client, {
            updateStockPrice: function (stock) {
                if (tableBinding == null) {
                    return;
                }
                var arr = [mapStockToArray(stock)];
                var td = new Office.TableData();
                td.rows = arr;
                tableBinding.setDataAsync(td, { startRow: stockRowMap[stock.Symbol] }, function(asyncResult) {
                    var x = 10;
                });
            },

            marketOpened: function() {
                $("#open").prop("disabled", true);
                $("#close").prop("disabled", false);
                $("#reset").prop("disabled", true);
            },

            marketClosed: function() {
                $("#open").prop("disabled", false);
                $("#close").prop("disabled", true);
                $("#reset").prop("disabled", false);
            },

            marketReset: function() {
                return initializeTable(tableBinding);
            }
        });

        // Start the connection
        $.connection.hub.start()
            //.pipe(init)
            .pipe(function () {
                return ticker.server.getMarketState();
            })
            .done(function (state) {
                if (state === 'Open') {
                    ticker.client.marketOpened();
                } else {
                    ticker.client.marketClosed();
                }

                // Wire up the buttons
                $("#open").click(function () {
                    ticker.server.openMarket();
                });

                $("#close").click(function () {
                    ticker.server.closeMarket();
                });

                $("#reset").click(function () {
                    ticker.server.reset();
                });
            });
    });
};
