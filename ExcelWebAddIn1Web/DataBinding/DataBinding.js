Office.onReady(() => {
    $(() => {
        app.initialize();

        $('#bind-to-existing-data').on('click', bindToExistingData);

        if (dataInsertionSupported()) {
            $('#insert-sample-data').show();
            $('#insert-sample-data').on('click', insertSampleData);
        }
    });
});
