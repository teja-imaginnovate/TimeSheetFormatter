# Time Sheet Formatter

A Java-based desktop application to convert raw time sheet CSV files into a standardized Excel report for easy reporting.

## Features

- Built-in GUI using Java Swing
- Load and preview raw time sheet CSV files
- Format and export time sheet data to Excel (`TimeSheet Report.xlsx`)

## Requirements

- JDK 17 or higher

## How to Use

1. **Build and Run the Application**
   - Compile and run the application using your preferred IDE or from the terminal:
     ```
     javac -d out src/**/*.java
     java -cp out com.imaginnovate.Main.java
     ```

2. **Open a Raw Time Sheet**
   - Click on the `Open CSV file` button in the GUI.
   - Select your raw time sheet CSV file.
   - The data will be loaded and displayed in a table.

3. **Format and Export**
   - Review the loaded data.
   - Click the `Format` button.
   - The application will process the data and generate a formatted Excel report named `TimeSheet Report.xlsx` in the same directory as the application.

## Output

- The formatted report will be saved as `TimeSheet Report.xlsx` in the application's directory.

## License

This project is for internal use only.
