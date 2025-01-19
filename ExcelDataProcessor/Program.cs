//Register  encoding provider
using ExcelDataReader;
using System.Data;
using System.Diagnostics;

// Register encoding provider
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

// Path to the Excel file
string folder = "RealEstatePricePrediction";
string filePath = $@"D:\Projects\{folder}\data\real_estate_data.xlsx";

if (!File.Exists(filePath))
{
  Console.WriteLine("File not found: " + filePath);
  return;
}

// Read the Excel file
using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
{
  using (var reader = ExcelReaderFactory.CreateReader(stream))
  {
    var dataSet = reader.AsDataSet();
    var table = dataSet.Tables[0]; //Assuming the first sheet

    // List to store cleaned data
    var cleanedData = new List<Dictionary<string, object>>();

    // Process rows
    for (int i = 1; i < table.Rows.Count; i++)// Skip header now
    {
      var price = table.Rows[i]["Price"]?.ToString();
      var roomCount = table.Rows[i]["Room_Count"]?.ToString();
      var grossSqM = table.Rows[i]["Gross_Square_Meters"]?.ToString();

      // Validate required fields
      if (string.IsNullOrEmpty(price) || string.IsNullOrEmpty(roomCount))
        continue; // Skip rows with missing data

      // Check for numeric validity
      if (!decimal.TryParse(price, out decimal parsedPrice) || parsedPrice <= 0)
        continue;

      // Store valid row
      cleanedData.Add(new Dictionary<string, object>
      {
        {"Price", parsedPrice },
        {"Room_Count", roomCount},
        {"Gross_Square_Meters", grossSqM}
      });
    }
    Console.WriteLine($"Cleaned Data Rows: {cleanedData.Count}");
  }
}