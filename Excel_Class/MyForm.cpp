#include "MyForm.h"
#include <msclr/marshal_cppstd.h>

using namespace ExcelClass;
using namespace Microsoft::Office;
using namespace Microsoft::Office::Core;

using namespace Microsoft::Office::Interop;
using namespace Microsoft::Office::Interop::Excel;
using namespace Microsoft::Office::Interop::PowerPoint;
using namespace std;


[STAThreadAttribute]

int main() {
	System::Windows::Forms::Application::Run(gcnew MyForm());
	return 0;
}

System::Void ExcelClass::MyForm::MyForm_Load(System::Object ^ sender, System::EventArgs ^ e)
{
	

	//int slide_firstIndex = 1;
	//String^ savePath = ".\\savePPT.pptx";
	//int slideHeight = 0;

	//PowerPoint::Application^ apt = gcnew PowerPoint::ApplicationClass();
	//PowerPoint::Presentations^ presen = apt->Presentations;

	//PowerPoint::Presentation^ presense1 = presen->Open(savePath, MsoTriState::msoFalse, MsoTriState::msoFalse, MsoTriState::msoTrue);
	//PowerPoint::Slide^ slide1 = presense1->Slides[1];
	//////�v���[���e�[�V�����V�K�쐬
	////PowerPoint::Presentation^ presense1 = presen->Add(MsoTriState::msoFalse);
	//////�X���C�h�ǉ�
	////PowerPoint::Slide^ slide1 = presense1->Slides->Add(slide_firstIndex, PowerPoint::PpSlideLayout::ppLayoutBlank);
	//width = (int)presense1->PageSetup->SlideWidth;
	//height = (int)presense1->PageSetup->SlideHeight;
	//int ct = 1;
	//

	//for each (Microsoft::Office::Interop::PowerPoint::Shape^ var in slide1->Shapes)
	//{
	//	if (var->HasTable == MsoTriState::msoTrue) {
	//		MessageBox::Show("table");
	//	}
	//	if (var->Type == MsoShapeType::msoEmbeddedOLEObject) {
	//		MessageBox::Show("umekomi");
	//		//���ߍ��݃G�N�Z���̋N��
	//		var->OLEFormat->DoVerb(1);
	//		//var->OLEFormat->DoVerb(2);
	//		//var->OLEFormat->Application;
	//		var->OLEFormat->Activate();
	//		Microsoft::Office::Interop::Excel::Workbook^ wb=(Microsoft::Office::Interop::Excel::Workbook^)var->OLEFormat->Object;
	//		Microsoft::Office::Interop::Excel::Worksheet^ worksheet = (Microsoft::Office::Interop::Excel::Worksheet^)wb->Worksheets[1];

	//		Microsoft::Office::Interop::Excel::Range^ testRange=(Microsoft::Office::Interop::Excel::Range^)worksheet->Cells[1,1];
	//		testRange->Value2 = "1";
	//		//MessageBox::Show(var->OLEFormat->Object->GetType()->ToString());
	//		//MessageBox::Show("table:" + var->HasTable.ToString());
	//		//MessageBox::Show("text:" + var->HasTextFrame.ToString());


	//	}
	//}

	//

	///*while (dataIndex < rowCount) {
	//	if (whileLoopEnd) {
	//		break;
	//	}
	//	createTable(dataIndex, slide1, height,ct);
	//	ct++;
	//}*/

	////�Z�[�u
	//presense1->SaveAs(savePath, Microsoft::Office::Interop::PowerPoint::PpSaveAsFileType::ppSaveAsDefault, MsoTriState::msoTrue);
	//
	////����
	//presense1->Close();
	//System::Runtime::InteropServices::Marshal::ReleaseComObject(presense1);
	//System::Runtime::InteropServices::Marshal::ReleaseComObject(presen);

	//apt->Quit();
	//System::Runtime::InteropServices::Marshal::ReleaseComObject(apt);

	//excel�̏�����
	Microsoft::Office::Interop::Excel::Application^ app_ = nullptr;
	Microsoft::Office::Interop::Excel::Workbooks^ wbs = nullptr;
	Microsoft::Office::Interop::Excel::Workbook^ workbook = nullptr;
	Microsoft::Office::Interop::Excel::Worksheets^ worksheets = nullptr;
	Microsoft::Office::Interop::Excel::Worksheet^ worksheet = nullptr;
	Microsoft::Office::Interop::Excel::Range^ testRange = nullptr;
	Microsoft::Office::Interop::Excel::ListObject^ lo = nullptr;

	////�J���t�@�C���̎w��
	String^ filePath = ".//sampleExcel__.xlsx";

	String^ saveFilePath = "C:\\Users\\chach\\Documents\\sampleExcel__123.xlsx";

	app_ = gcnew Microsoft::Office::Interop::Excel::ApplicationClass();
	////Excel�u�b�N�̕\���͂��Ȃ�
	app_->Visible = false;
	app_->DisplayAlerts = false;

	wbs = app_->Workbooks;
	//�V�K�ǉ�
	workbook=wbs->Add(Type::Missing);
	//�ۑ�
	workbook->SaveAs(saveFilePath, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Excel::XlSaveAsAccessMode::xlNoChange, Excel::XlSaveConflictResolution::xlOtherSessionChanges, Type::Missing, Type::Missing, Type::Missing, Type::Missing);
	
	if (lo != nullptr)
	{
		System::Runtime::InteropServices::Marshal::ReleaseComObject(lo);
		lo = nullptr;
	}
	GC::Collect();
	GC::WaitForPendingFinalizers();
	GC::Collect();
	if (testRange != nullptr)
	{
		System::Runtime::InteropServices::Marshal::ReleaseComObject(testRange);
		testRange = nullptr;
	}
	GC::Collect();
	GC::WaitForPendingFinalizers();
	GC::Collect();
	workbook->Close(Type::Missing, Type::Missing, Type::Missing);
	if (workbook != nullptr)
	{
		System::Runtime::InteropServices::Marshal::ReleaseComObject(workbook);
		workbook = nullptr;
	}
	GC::Collect();
	GC::WaitForPendingFinalizers();
	GC::Collect();
	if (wbs != nullptr)
	{
		System::Runtime::InteropServices::Marshal::ReleaseComObject(wbs);
		wbs = nullptr;
	}
	GC::Collect();
	GC::WaitForPendingFinalizers();
	GC::Collect();

	if (app_ != nullptr)
	{
		app_->Quit();
		System::Runtime::InteropServices::Marshal::ReleaseComObject(app_);
		app_ = nullptr;
	}
	GC::Collect();
	GC::WaitForPendingFinalizers();
	GC::Collect();

	app_ = gcnew Microsoft::Office::Interop::Excel::ApplicationClass();
	////Excel�u�b�N�̕\���͂��Ȃ�
	app_->Visible = false;
	app_->DisplayAlerts = false;

	////�t�@�C���p�X����u�b�N���J��
	workbook = (Microsoft::Office::Interop::Excel::Workbook^)(app_->Workbooks->Open(
		saveFilePath,
		Type::Missing,
		Type::Missing,
		Type::Missing,
		Type::Missing,
		Type::Missing,
		Type::Missing,
		Type::Missing,
		Type::Missing,
		Type::Missing,
		Type::Missing,
		Type::Missing,
		Type::Missing,
		Type::Missing,
		Type::Missing));

	//worksheets = (Microsoft::Office::Interop::Excel::Worksheets^)workbook->Sheets;
	worksheet = (Microsoft::Office::Interop::Excel::Worksheet^)workbook->Sheets->Add(Type::Missing,Type::Missing,1,Type::Missing);

	//�ۑ�
	workbook->SaveAs(saveFilePath, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Excel::XlSaveAsAccessMode::xlNoChange, Type::Missing, Type::Missing, Type::Missing, Type::Missing, Type::Missing);


	////�ۑ�
	//workbook->SaveCopyAs(saveFilePath);

	//workbook->Close(Type::Missing, Type::Missing, Type::Missing);

	////�ꖇ�ڂ̃��[�N�V�[�g���J��
	//worksheet = (Microsoft::Office::Interop::Excel::Worksheet^)workbook->Worksheets[1];

	//testRange=(Microsoft::Office::Interop::Excel::Range^)worksheet->Cells[1,1];
	////std::string dam = "1,2,3,4,6";
	//String^ sam = "1,2,3,4,5";
	//sam = sam->Replace(",", ",\n");
	////sam = sam->Substring(0,sam->Length-1);
	////dam=dam.replace(dam.begin(), dam.end(), ",", ",\n");
	////MessageBox::Show(sam);
	////cli::array<String^>^ arr = sam->Split(',');

	////testRange->Value2 = sam;

	////���[�N�V�[�g���̃��X�g�I�u�W�F�N�g���擾
	//System::Collections::IEnumerator^ enu=worksheet->ListObjects->GetEnumerator();
	//int beforeData = 1;
	//int labelNum = 1;
	//int afterData = 2;
	////enum�𔽕�����
	//while (enu->MoveNext()) {
	//	//�e�[�u�����擾
	//	lo = (Excel::ListObject^)enu->Current;
	//	//�J�n�ʒu��0�ł͂Ȃ��A1����J�n����
	//	testRange = (Excel::Range^)lo->Range->Cells[labelNum, beforeData];
	//	//testRange->RowHeight = 5;
	//	//�e�[�u���ɗ��ǉ�
	//	lo->ListColumns->Add(2);
	//	lo->ListColumns->Add(3);

	//	testRange= (Excel::Range^)lo->Range->Cells[2, 2];
	//	//MessageBox::Show(msclr::interop::marshal_as<System::String^>(dam));
	//	testRange->Value2 = sam;
	//	
	//	//�w�b�_�[�������A�ŏ��̍s����J�n
	////	for (int i = 2; i < lo->ListRows->Count + 2; i++) {
	////		//���H�Ώۃf�[�^���擾
	////		testRange = (Excel::Range^)lo->Range->Cells[i, beforeData];
	////		String^ tmpData = testRange->Text->ToString();
	////		//4���ڂɋ�؂蕶����}��
	////		tmpData = tmpData->Insert(4, "#");
	////		//��؂蕶���̑O��ŕ�����
	////		cli::array<String^>^ arr = tmpData->Split('#');
	////		//��؂蕶���̑O����2��ڂɑ��
	////		lo->Range->Cells[i, 2] = arr[0];
	////		//��؂蕶���̌㔼��3��ڂɑ��
	////		lo->Range->Cells[i, 3] = arr[1];

	//	}
	////	//�\�[�g�O�Ƀ\�[�g�������N���A���Ă���
	////	lo->Sort->SortFields->Clear();
	////	//Range��2��ڂ��w��
	////	testRange = (Excel::Range^)lo->ListColumns[2]->Range;
	////	try {
	////		//�ŗD��ł���2��ڂ̃\�[�g�������w��
	////		lo->Sort->SortFields->Add2(
	////			testRange, Excel::XlSortOn::xlSortOnValues, Excel::XlSortOrder::xlAscending, Type::Missing, Excel::XlSortDataOption::xlSortNormal, Type::Missing);
	////	}catch(Exception^ e) {
	////		MessageBox::Show(e->ToString());
	////	}
	////	//range��3��ڂ��w��
	////	testRange = (Excel::Range^)lo->ListColumns[3]->Range;
	////	//2�ԖڂɗD�悷��3��ڂ̃\�[�g�������w��
	////	lo->Sort->SortFields->Add2(
	////		testRange, Excel::XlSortOn::xlSortOnValues, Excel::XlSortOrder::xlAscending, Type::Missing, Excel::XlSortDataOption::xlSortNormal, Type::Missing);
	////	//�\�[�g���s�̂��߂̃v���p�e�B��ݒ�
	////	lo->Sort->Header = XlYesNoGuess::xlYes;
	////	lo->Sort->MatchCase = false;
	////	lo->Sort->Orientation = XlSortOrientation::xlSortColumns;
	////	lo->Sort->SortMethod = XlSortMethod::xlPinYin;
	////	lo->Sort->Apply();

	////	break;
	////}
	////lo->ListColumns[3]->Delete();
	////lo->ListColumns[2]->Delete();


	//workbook->Save();

	////Excel�̃v���Z�X����鏈��
	if (lo != nullptr)
	{
		System::Runtime::InteropServices::Marshal::ReleaseComObject(lo);
		lo = nullptr;
	}
	GC::Collect();
	GC::WaitForPendingFinalizers();
	GC::Collect();
	if (testRange != nullptr)
	{
		System::Runtime::InteropServices::Marshal::ReleaseComObject(testRange);
		testRange = nullptr;
	}
	GC::Collect();
	GC::WaitForPendingFinalizers();
	GC::Collect();
	if (worksheet != nullptr)
	{
		System::Runtime::InteropServices::Marshal::ReleaseComObject(worksheet);
		worksheet = nullptr;
	}
	GC::Collect();
	GC::WaitForPendingFinalizers();
	GC::Collect();
	if (worksheets != nullptr)
	{
		System::Runtime::InteropServices::Marshal::ReleaseComObject(worksheets);
		worksheets = nullptr;
	}
	GC::Collect();
	GC::WaitForPendingFinalizers();
	GC::Collect();

	workbook->Close(Type::Missing, Type::Missing, Type::Missing);
	if (workbook != nullptr)
	{
		System::Runtime::InteropServices::Marshal::ReleaseComObject(workbook);
		workbook = nullptr;
	}
	GC::Collect();
	GC::WaitForPendingFinalizers();
	GC::Collect();
	if (wbs != nullptr)
	{
		System::Runtime::InteropServices::Marshal::ReleaseComObject(wbs);
		wbs = nullptr;
	}
	GC::Collect();
	GC::WaitForPendingFinalizers();
	GC::Collect();

	if (app_ != nullptr)
	{
		app_->Quit();
		System::Runtime::InteropServices::Marshal::ReleaseComObject(app_);
		app_ = nullptr;
	}
	GC::Collect();
	GC::WaitForPendingFinalizers();
	GC::Collect();

	//return System::Void();
}

System::Void ExcelClass::MyForm::createTable(int index, PowerPoint::Slide^ slide1, int height,int tableNum)
{
	
	PowerPoint::Shape^ containTable = slide1->Shapes->AddTable(10, 4, (350*(tableNum-1)), startY_value, 300, 300);
	PowerPoint::Table^ table = containTable->Table;
	//�\�̏����ʒu
	int totalHeight = startY_value;
	int tmpHeight = 0;

	int FontSize = 10;
	//int rowCount = 50;
	for (int j = 1; j < rowCount; j++) {
		
		if (totalRowCount >= rowCount) {
			table->Rows[table->Rows->Count]->Delete();
			whileLoopEnd = true;
			break;
		}

		totalHeight += tmpHeight;
		tmpHeight = 0;

		if (height <= totalHeight) {
			table->Rows[table->Rows->Count]->Delete();
			dataIndex += table->Rows->Count;
			//MessageBox::Show(dataIndex.ToString());
			break;
		}
		if (j > 9) {
			table->Rows->Add(j);
		}
		totalRowCount++;
		for (int i = 1; i < table->Columns->Count + 1; i++) {

			table->Rows[j]->Cells[i]->Shape->TextFrame->TextRange->Font->Size = FontSize;
			table->Rows[j]->Cells[i]->Shape->TextFrame->TextRange->Text = j.ToString();

			//�p���[�|�C���g���̃Z���̃e�L�X�g�𒆉��񂹂ɂ���
			table->Rows[j]->Cells[i]->Shape->TextFrame->HorizontalAnchor = MsoHorizontalAnchor::msoAnchorCenter;
			//�Ȃ��A���l�╶����Ȃǂ̏����̐ݒ�̓G�N�Z���@�\�ł̑Ή��Ȃ̂Ńp���[�|�C���g�ł͎w�肵�悤���Ȃ�
			if (i == 1 && j == 1) {
				//���s�̋�؂蕶���̃e�X�g
				table->Rows[j]->Cells[i]->Shape->TextFrame->TextRange->Text = "100,200,300,400" + "\r\n" + "50,60,70,80";
				table->Rows[j]->Cells[i]->Shape->TextFrame->TextRange->Font->Size = FontSize;
				//�t�H���g�T�C�Y��10�{�ɗ�̕���ݒ�
				table->Columns[i]->Width = FontSize * 10;
				//totalHeight += table->Rows[j]->Cells[i]->Shape->Height;
			}
			if (i == 2 && j == 2) {
				table->Rows[j]->Cells[i]->Shape->TextFrame->TextRange->Text = "sampleTest";
				table->Rows[j]->Cells[i]->Shape->TextFrame->TextRange->Font->Size = FontSize;
				//�t�H���g�T�C�Y��10�{�ɗ�̕���ݒ�
				table->Columns[i]->Width = FontSize * 10;
				totalHeight += table->Rows[j]->Cells[i]->Shape->Height;
			}
			//�Z���̍����̍X�V
			if (tmpHeight < table->Rows[j]->Cells[i]->Shape->Height) {
				tmpHeight = table->Rows[j]->Cells[i]->Shape->Height;
			}
		}
	}
	return System::Void();
}
