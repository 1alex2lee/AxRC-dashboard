<section class="${styles.amrcDashboard} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
    <div>
      <h1>Project Portfolio Dashboard</h1>
      <button onclick=open_file()>Open and Read file</button>
      <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/jszip.js"></script>
      <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/xlsx.js"></script>
      <script>
        open_file = function() {
          this.parseExcel = function(file) {
            var reader = new FileReader();
        
            reader.onload = function(e) {
              var data = e.target.result;
              var workbook = XLSX.read(data, {
                type: 'binary'
              });
              workbook.SheetNames.forEach(function(sheetName) {
                // Here is your object
                var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
                var json_object = JSON.stringify(XL_row_object);
                console.log(json_object);
              })
            };
            reader.onerror = function(ex) {
              console.log(ex);
            };
            reader.readAsBinaryString(file);
          };
        }
      </script>
    </div>
  </section>`;
  