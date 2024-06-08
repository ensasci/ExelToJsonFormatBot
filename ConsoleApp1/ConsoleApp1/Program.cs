using OfficeOpenXml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;

class Program
{
	// Gerekli kütüphaneler
	//NewtonsoftJson
	//EEPlus
	static void Main(string[] args)
	{

		
		// Excel dosyasının yolunu belirtin
		string excelFilePath = "C:\\Users\\Enes\\Desktop\\Classeur1.xlsx";
		//telefonlar adında bir string listesi oluşturdum bunu daha sonrasında aynı telefon numaralarını kayıt etmemek için kullanıcam
		List<string> phones = new List<string>();
		//Jobject türünde bir liste oluştrduk 
		List<JObject> jsonArray = new List<JObject>();
		//file info ile excel dosa yolunu verdiğim dosyayı açıyorum
		FileInfo fileInfo = new FileInfo(excelFilePath);
		using (ExcelPackage package = new ExcelPackage(fileInfo))
		{
				ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
				//satır sayısını bulduk
				int rowCount = worksheet.Dimension.Rows;
			 //kayıtlarım 2.satırdan itibaren başladığı için 2 de başlatıp uzunluk kadar bir for döngüsü oluşturduk
				for (int row = 2; row <= rowCount; row++)
				{
				// row değerinin 1 sütündeki textini ve 2 sütünda bulunan texti fulnname atadık
					string fullName = worksheet.Cells[row, 1].Text + " " + worksheet.Cells[row, 2].Text; 
					string phoneNumber = worksheet.Cells[row, 3].Text; // PhoneNumber sütunu
					string emailAddress = "";
					var phone = ""; ;
				for (int i = 0; i < phoneNumber.Length; i++)
				{
					//telefon numarası içindeki istenmeyen karakterli silmek için phonenumberın uzunluğu kadar döndük ve her bir karakterine baktık 48 57 asciii kodlarındaki değerleri temsil ediyor
					//char int türüne döndüğü için bu şekilde ascii kodları ile bir kontrol yaptık ve 0 ile 9 arasındaki tüm ifadeleri kaldırdık (boşkuk,( , ) , - ) gibi ifaedeleri
				 if (phoneNumber[i] >= 48 && phoneNumber[i] <= 57)
					{
						phone += phoneNumber[i].ToString();
					}
				}
				//sonra bu phones listemde var mı diye kontrol ediyorum aynı numarayı bularak aynı kayıdı tekrar eklemesin diye eğer yoksa 
				if (!phones.Contains(phone))
				{
					//listeme ekle
					phones.Add(phone);
					//email adres formatını telefonnumarası@mkparisconcept.com şeklinde ayarla
					emailAddress = $"{phone}@mkparisconcept.com";
					//json objesini oluşrur
					JObject jsonObject = new JObject
				{
					{ "CreatedAt", new JObject { { "$date", DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ") } } },
					{ "ModifiedAt", null },
					{ "RecordStatus", 1 },
					{ "CreatingUserId", null },
					{ "UpdatingUserId", null },
					{ "TenantId", "664d0954cd65132ff0928cbf" },
					{ "FullName", fullName },
					{ "EmailAddress", emailAddress },
					{ "PhoneNumber", phone }
				};
					// json listeme jsonobjemi ekle
					jsonArray.Add(jsonObject);
				}
			}
		}

		// JSON array'i string formatında oluştur
		string jsonOutput = JsonConvert.SerializeObject(jsonArray, Formatting.Indented);

		// JSON'u bir dosyaya yazdır
		File.WriteAllText("output.json", jsonOutput);
		Console.WriteLine("JSON oluşturuldu ve 'output.json' dosyasına yazıldı.");
	}
}