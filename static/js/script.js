document.addEventListener('DOMContentLoaded', function() {
    var data = [
        ['Flight', 'Date', 'ST', 'State', 'STD', 'STA', 'Best DT', 'Best AT', 'From', 'To', 'Reg.', 'Own / Sub', 'Delay', 'Pax(F/C/Y)'],
        ['','','','','','','','','','','','','','']
    ];

    var container = document.getElementById('table');
    var hot = new Handsontable(container, {
        data: data,
        rowHeaders: true,
        colHeaders: true,
        contextMenu: true,
        minSpareRows: 1,
        columns: [
            {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {}
        ]
    });

    document.getElementById('dataForm').addEventListener('submit', function() {
        var table_data = hot.getData();
        document.getElementById('table_data').value = JSON.stringify(table_data);
    });
});
