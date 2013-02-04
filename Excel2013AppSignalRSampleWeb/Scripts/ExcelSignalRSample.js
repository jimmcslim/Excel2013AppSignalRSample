// This function is run when the app is ready to start interacting with the host application
// It ensures the DOM is ready before adding click handlers to buttons
Office.initialize = function (reason) {
    $(document).ready(function () {

        var stockProperties = ['Symbol', 'Price', 'DayHigh', 'DayLow', 'DayOpen', 'Change', 'LastChange', 'PercentChange'];
        var stockRowMap = {};
        var tableBinding;

        $('#stockProperties').append(stockProperties.join(', '));

        $('#clearBinding').prop('disabled', true);

        $('#setBinding').click(function () {

            $('.alert-binding').alert('close');
            
            // It would be nice if restrictions on the Binding parameters could be optionally specified here,
            // such as the 'shape' of a TableBinding/MatrixBinding (i.e. min/max number of rows/columns), or the
            // expected headers of a table's columns. As it is, we will check that here...
            Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Table,
                { id: 'StockTable', promptText: 'Select a table to bind to.' },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        $('"#messages').prepend(
                            '<div class="alert alert-error alert-binding">' +
                                '<button type="button" class="close" data-dismiss="alert">&times;</button>' +
                                'An error occurred attempting to set the TableBinding: ' + asyncResult.error.message +
                            '</div>');
                        return;
                    }

                    tableBinding = asyncResult.value;
                    
                    // In Excel 2013, it seems you can't create a table without headers, so a call to binding.hasHeaders
                    // isn't necessary, we can just get the headers async.
                    tableBinding.getDataAsync({ rowCount: 0 }, function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                            $('#messages').prepend(
                                '<div class="alert alert-error alert-binding">' +
                                    '<button type="button" class="close" data-dismiss="alert">&times;</button>' +
                                    'An error occurred attempting to get the table headers: ' + asyncResult.error.message +
                                '</div>');
                            Office.context.document.bindings.releaseByIdAsync('StockTable');
                            return;
                        }

                        var headers = asyncResult.value.headers[0];
                        
                        // Determine if there are any unrecognised headers...
                        var unrecognisedHeaders = headers.filter(function (element) {
                            return stockProperties.indexOf(element) < 0;
                        });
                        
                        // and if so... create an error message and release the binding.
                        if (unrecognisedHeaders.length !== 0) {
                            $('#messages').prepend(
                                '<div class="alert alert-error alert-binding">' +
                                    '<button type="button" class="close" data-dismiss="alert">&times;</button>' +
                                    'There are unrecognised column headers; please remove and create the binding again: ' +
                                    unrecognisedHeaders.join(',') +
                                '</div>');

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

                        $('#messages').prepend(
                            '<div class="alert alert-success alert-binding">' +
                                '<button type="button" class="close" data-dismiss="alert">&times;</button>' +
                                'Binding successful.' +
                            '</div>'
                        );

                        $('#setBinding').prop('disabled', true);
                        $('#clearBinding').prop('disabled', false);
                    });
                });
        });

        $('#clearBinding').click(function () {
            $('#setBinding').prop('disabled', false);
            $('#clearBinding').prop('disabled', true);
            Office.context.document.bindings.releaseByIdAsync('StockTable', function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    $('#messages').prepend(
                        '<div class="alert alert-error alert-binding">' +
                            '<button type="button" class="close" data-dismiss="alert">&times;</button>' +
                            'An error occurred attempting to clear the binding: ' + asyncResult.error.message +
                        '</div>'
                    );
                }
            });
        });
        
        function initializeTable(binding) {
            ticker.server.getAllStocks().done(function(stocks) {
                $(stocks).each(function(idx, val) {
                    stockRowMap[val.Symbol] = idx;
                });
                stocks = stocks.map(mapStockToArray);

                binding.addRowsAsync(stocks, function(asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {

                    }
                });
            });
        }
        
        var ticker = $.connection.stockTicker;

        function mapStockToArray() {
            $('#messages').append('<p>mapStockToArray function has not yet been defined.</p>');
        }

        $.extend(ticker.client, {
            updateStockPrice: function (stock) {
                if (tableBinding === null) {
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
                $('#open').prop('disabled', true);
                $('#close').prop('disabled', false);
                $('#reset').prop('disabled', true);
            },

            marketClosed: function() {
                $('#open').prop('disabled', false);
                $('#close').prop('disabled', true);
                $('#reset').prop('disabled', false);
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
                $('#open').click(function () {
                    ticker.server.openMarket();
                });

                $('#close').click(function () {
                    ticker.server.closeMarket();
                });

                $('#reset').click(function () {
                    ticker.server.reset();
                });
            });
    });
};
