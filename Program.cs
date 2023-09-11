using ExcelFileWriter;

class Program
{
    private const string saveDialogMessage = "Excel files (*.xlsx)|*.xlsx";
    private const string saveDialogExtensions = "xlsx";

    [STAThread]
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
            string path = GetPathFromSaveFileDialog();

            if (string.IsNullOrEmpty(path))
            {
                MessageBox.Show("Не выбрано имя файла!");
            }
            else
            {
                ExcelDataWriter writer = new();
                writer.Write(cars, path);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.StackTrace);
        }
    }

    /// <summary>
    /// Вызов стандартного окна сохранения файла Windows.
    /// Значение поля OverwritePrompt устанавливается в false для избежания "дублирования" запроса о замене файла - стандартным окном Windows и окном Excel
    /// </summary>
    /// <returns></returns>
    private static string GetPathFromSaveFileDialog()
    {
        SaveFileDialog saveFileDialog = new SaveFileDialog();
        saveFileDialog.Filter = saveDialogMessage;
        saveFileDialog.DefaultExt = saveDialogExtensions;
        saveFileDialog.AddExtension = true;
        saveFileDialog.OverwritePrompt = false;
        saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        if (saveFileDialog.ShowDialog() == DialogResult.OK)
        {
            return saveFileDialog.FileName;
        }
        else
        {
            return string.Empty;
        }
    }
}