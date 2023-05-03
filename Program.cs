using System.Web;
using System.Net;
using OfficeOpenXml;
public class Program{
    private static string[] bloods = { "A+", "A-", "B+", "B-", "AB+", "AB-", "O+", "O-" };

    private static void Main(string[] args){
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var FILE = new FileInfo(@"files/Data.xlsx");
        var DATA = generateData(2500);
        SaveExcelFile(FILE, DATA);
    }

    private static List<Row> generateData(int i){
        List<Row> data = new List<Row>();
        int p = (0/i)*100, n = (0/i)*100;
        for(int j = 0; j < i; j++){
            data.Add(generateFakeData());
            n = (int)(((j+1)/(double)i)*100);
            if(p != n){
                Console.WriteLine(n + "%");
                p = n;
            }
        }
        return data;
    }
    private static Row generateFakeData(){
        var faker = new Bogus.Faker("ar");
        var fakerEN = new Bogus.Faker();

        var Name = faker.Name;
        var Date = faker.Date;
        var Gender = faker.Person.Gender;
        var fName = Name.FirstName(Gender);
        var lName = Name.LastName();

        int EMPINTERGRATIONID = faker.Random.Number(1, 9999999); // 7
        string EMPCARDID = faker.Random.AlphaNumeric(20).ToUpper(); // 20 ?
        var EMPNATIONALID = faker.Random.Number(2000000000);
        int DEPID = faker.Random.Number(1, 7);
        string CATID = faker.Random.Number(1, 1000000000).ToString(); // 10
        string EMPNAMEAR = fName + " " + lName; // 83
        string EMPNAMEEN = translate(fName, "ar", "en") + " " + translate(lName, "ar", "en"); // 83
        var EMPJOINDATE = Date.Past(faker.Random.Int(0, 5));
        var EMPLEAVEDATE = Date.Between(EMPJOINDATE, DateTime.Now);
        int EMPSTATUS = faker.Random.Number(0, 1); // ?
        string EMPPWD = faker.Random.Number(0, 1).ToString(); // ?
        string EMPEMAILID = faker.Internet.Email(fName, lName);
        string EMPMOBILENO = "+9665" + faker.Random.Number(99999999).ToString(); //0598585594
        int NATID = faker.Random.Number(1, 100); // 3 ?
        string JOBTITLE = fakerEN.Name.JobTitle(); ///
        string EMPJOBTITLEAR = translate(JOBTITLE, "en", "ar");
        string EMPJOBTITLEEN = translate(JOBTITLE, "ar", "en");
        string EMPBLOOD = faker.Random.ArrayElement(bloods);
        char EMPGENDER = Gender.ToString()[0];
        string EMPDESCRIPITON = fakerEN.Lorem.Sentence();
        string EMPDOMAINUSERNAME = faker.Internet.UserName(fName, lName);

        Row data = new Row();
        data.EMPINTERGRATIONID = EMPINTERGRATIONID;
        data.EMPCARDID = EMPCARDID;
        data.EMPNATIONALID = EMPNATIONALID;
        data.DEPID = DEPID;
        data.CATID = CATID;
        data.EMPNAMEAR = EMPNAMEAR;
        data.EMPNAMEEN = EMPNAMEEN;
        data.EMPJOINDATE = EMPJOINDATE;
        data.EMPLEAVEDATE = EMPLEAVEDATE;
        data.EMPSTATUS = EMPSTATUS;
        data.EMPPWD = EMPPWD;
        data.EMPEMAILID = EMPEMAILID;
        data.EMPMOBILENO = EMPMOBILENO;
        data.NATID = NATID;
        data.EMPJOBTITLEAR = EMPJOBTITLEAR;
        data.EMPJOBTITLEEN = EMPJOBTITLEEN;
        data.EMPBLOOD = EMPBLOOD;
        data.EMPGENDER = EMPGENDER;
        data.EMPDESCRIPITON = EMPDESCRIPITON;
        data.EMPDOMAINUSERNAME = EMPDOMAINUSERNAME;
        return data;
    }
    private static string translate(string input, string from, string to){
        var fromLanguage = from;
        var toLanguage = to;
        var url = $"https://translate.googleapis.com/translate_a/single?client=gtx&sl={fromLanguage}&tl={toLanguage}&dt=t&q={HttpUtility.UrlEncode(input)}";
        var webclient = new WebClient{
            Encoding = System.Text.Encoding.UTF8
        };
        var result = webclient.DownloadString(url);
        try{
            result = result.Substring(4, result.IndexOf("\"", 4, StringComparison.Ordinal) - 4);
            return result;
        }
        catch (Exception e1){

            return "error" + e1;
        }
    }  
    private static void SaveExcelFile(FileInfo file, List<Row> data){
        if (file.Exists){
            file.Delete();
        }
        using var package = new ExcelPackage(file);
        var ws = package.Workbook.Worksheets.Add("MainReport");
        var range = ws.Cells["A1"].LoadFromCollection(data, true);

        range.AutoFitColumns();
        package.Save();
    } 
}

public class Row{
    public int EMPINTERGRATIONID {get; set;}
    public string EMPCARDID {get; set;} // 20 ?
    public int EMPNATIONALID {get; set;}
    public int DEPID {get; set;}
    public string CATID {get; set;} // 10
    public string EMPNAMEAR {get; set;} // 83
    public string EMPNAMEEN {get; set;} // 83
    public DateTime EMPJOINDATE {get; set;}
    public DateTime EMPLEAVEDATE {get; set;}
    public int EMPSTATUS {get; set;} // ?
    public string EMPPWD {get; set;} // ?
    public string EMPEMAILID {get; set;}
    public string EMPMOBILENO {get; set;}
    public int NATID {get; set;} // 3 ?
    public string EMPJOBTITLEAR {get; set;}
    public string EMPJOBTITLEEN {get; set;}            
    public string EMPBLOOD {get; set;}
    public char EMPGENDER {get; set;}
    public string EMPDESCRIPITON {get; set;}
    public string EMPDOMAINUSERNAME {get; set;}

}