const XLSX = require('xlsx');
const readline = require('readline');
const fs = require('fs');

// Create readline interface
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

// Get the file path from user input
rl.question('Enter the path to the Excel file: ', (filePath) => { 
  // Checking if file exists
  if (!fs.existsSync(filePath)) {
    console.log('File not found.');
    rl.close();
    return;
  }

  console.log('Reading file...');

  // Reading file
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  console.log('File read successfully.');

  // Extract data from the sheet
  const data = XLSX.utils.sheet_to_json(worksheet);

  // User input for column name
  rl.question('Enter the name of the column you want to search: ', (columnName) => {
    // Check if column name exists 
    if (!data[0].hasOwnProperty(columnName)) {
      console.log('Error: The provided column name does not exist in the file.');
      rl.close();
      return;
    }

    // Value to be replaced
    rl.question(`Enter the value you want to replace in column "${columnName}": `, (valueToReplace) => {
      // Check if value exists in that column
      const valueExists = data.some((row) => String(row[columnName]) === String(valueToReplace));
      if (!valueExists) {
        console.log(`Error: The provided value "${valueToReplace}" does not exist in the column "${columnName}".`);
        rl.close();
        return;
      }

      // New value
      rl.question('Enter the new value: ', (newValue) => {
        console.log('Replacing values...');

        // Replace value in that column
        data.forEach((row) => {
          if (String(row[columnName]) === String(valueToReplace)) {
            row[columnName] = newValue; 
          }
        });

        console.log('Values replaced.');

        // Convert data back to worksheet format
        const newWorksheet = XLSX.utils.json_to_sheet(data);
        workbook.Sheets[sheetName] = newWorksheet;

        // Write the updated workbook back
        console.log('Writing updated file...');
        XLSX.writeFile(workbook, filePath);
        console.log('Replacement done!');
        
        rl.close();
      });
    });
  }); 
});
 