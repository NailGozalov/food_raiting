document.getElementById('downloadExcel').addEventListener('click', function () {
    
   
    var wb = XLSX.utils.book_new();
    
    
    var table = document.getElementById('resultsTable');
    var rows = Array.from(table.querySelectorAll('tbody tr'));
    var data = rows.map(row => Array.from(row.cells).slice(1).map(cell => cell.textContent));
    
   
    var headers = ['Soup', 'Salad', 'Main course', 'Compote', 'Garnish', 'Date & Time'];
    

    var ws = XLSX.utils.book_new();
    
 
    XLSX.utils.sheet_add_aoa(ws, [headers, ...data]);
    
   
    var headerStyle = { font: { bold: true } };
    headers.forEach((header, i) => {
        var cellAddress = XLSX.utils.encode_cell({ c: i, r: 0 }); 
        ws[cellAddress].s = headerStyle; 
    });
    
    
    var columnWidths = [15, 15, 15, 15, 15, 20]; 
    columnWidths.forEach((width, i) => {
        var colAddress = XLSX.utils.encode_col(i);
        ws['!cols'] = ws['!cols'] || [];
        ws['!cols'].push({ wch: width });
    });
    
   
    XLSX.utils.book_append_sheet(wb, ws, "Results");
    

    XLSX.writeFile(wb, 'Results.xlsx');
});



