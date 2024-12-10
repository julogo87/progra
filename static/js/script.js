document.addEventListener('DOMContentLoaded', function() {
    var container = document.getElementById('table');
    var hot = new Handsontable(container, {
        data: [],
        rowHeaders: true,
        colHeaders: [
            'N', 'Flight', 'Date', 'ST', 'State', 'STD', 'STA', 'Best DT', 'Best AT', 
            'From', 'To', 'Reg.', 'Own / Sub', 'Delay', 'Pax(F/C/Y)', 'Crew', 'Notas', 'Tripadi'
        ],
        contextMenu: true,
        minSpareRows: 1,
        columns: [
            {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, 
            { data: 'crew', type: 'text' },    // Columna Crew
            { data: 'notas', type: 'text' },   // Columna Notas
            { data: 'tripadi', type: 'text' }  // Columna Tripadi
        ]
    });

    document.getElementById('dataForm').addEventListener('submit', function() {
        var table_data = hot.getData();
        var headers = hot.getColHeader();
        var json_data = table_data.map(row => {
            var rowData = {};
            headers.forEach((header, index) => {
                rowData[header] = row[index];
            });
            return rowData;
        });
        document.getElementById('table_data').value = JSON.stringify(json_data);
    });
});
