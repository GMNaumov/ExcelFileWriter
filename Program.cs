using ExcelFileWriter;

class Program
{
    static void Main(string[] args)
    {
        Car carOne = new() { Brand = "BMW", Model = "M5", Price = 2_500_000.0 };
        Car carTwo = new() { Brand = "Mercedes", Model = "E250", Price = 3_700_000.0 };
        Car carThree = new() { Brand = "Audi", Model = "A6", Price = 2_200_000.0 };

        List<Car> cars = new()
        {
            carOne,
            carTwo,
            carThree
        };

        try
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Application.StartupPath;
            StreamWriter streamWriter;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                streamWriter = new StreamWriter(saveFileDialog.FileName);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.StackTrace);
        }

        ExcelDataWriter writer = new();
        string path = @"D:\NewFile.xlsx";
        writer.Write(cars, path);
    }
}