#include "MyForm.h"

using namespace ExcelClass;
using namespace Microsoft::Office;
using namespace Microsoft::Office::Core;

using namespace Microsoft::Office::Interop;
using namespace Microsoft::Office::Interop::Excel;
using namespace Microsoft::Office::Interop::PowerPoint;


[STAThreadAttribute]

int main() {
	System::Windows::Forms::Application::Run(gcnew MyForm());
	return 0;
}

System::Void ExcelClass::MyForm::MyForm_Load(System::Object ^ sender, System::EventArgs ^ e)
{
	

	int slide_firstIndex = 1;
	String^ savePath = ".\\savePPT.pptx";
	int slideHeight = 0;

	PowerPoint::Application^ apt = gcnew PowerPoint::ApplicationClass();
	PowerPoint::Presentations^ presen = apt->Presentations;
	//プレゼンテーション新規作成
	PowerPoint::Presentation^ presense1 = presen->Add(MsoTriState::msoFalse);
	//スライド追加
	PowerPoint::Slide^ slide1 = presense1->Slides->Add(slide_firstIndex, PowerPoint::PpSlideLayout::ppLayoutBlank);
	width = (int)presense1->PageSetup->SlideWidth;
	height = (int)presense1->PageSetup->SlideHeight;
	int ct = 1;
	while (dataIndex < rowCount) {
		if (whileLoopEnd) {
			break;
		}
		createTable(dataIndex, slide1, height,ct);
		ct++;
	}

	//セーブ
	presense1->SaveAs(savePath, Microsoft::Office::Interop::PowerPoint::PpSaveAsFileType::ppSaveAsDefault, MsoTriState::msoTrue);
	
	//閉じる
	presense1->Close();
	System::Runtime::InteropServices::Marshal::ReleaseComObject(presense1);
	System::Runtime::InteropServices::Marshal::ReleaseComObject(presen);

	apt->Quit();
	System::Runtime::InteropServices::Marshal::ReleaseComObject(apt);

	//excelの初期化
	Microsoft::Office::Interop::Excel::Application^ app_ = nullptr;
	Microsoft::Office::Interop::Excel::Workbook^ workbook = nullptr;
	Microsoft::Office::Interop::Excel::Worksheet^ worksheet = nullptr;
	Microsoft::Office::Interop::Excel::Range^ testRange = nullptr;
	Microsoft::Office::Interop::Excel::ListObject^ lo = nullptr;

	//開くファイルの指定
	String^ filePath = ".//sampleExcel.xlsx";

	app_ = gcnew Microsoft::Office::Interop::Excel::ApplicationClass();
	//Excelブックの表示はしない
	app_->Visible = false;
	//ファイルパスからブックを開く
	workbook = (Microsoft::Office::Interop::Excel::Workbook^)(app_->Workbooks->Open(
		filePath,
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
	//一枚目のワークシートを開く
	worksheet = (Microsoft::Office::Interop::Excel::Worksheet^)workbook->Worksheets[1];
	//ワークシート内のリストオブジェクトを取得
	System::Collections::IEnumerator^ enu=worksheet->ListObjects->GetEnumerator();
	int beforeData = 1;
	int labelNum = 1;
	int afterData = 2;
	//enumを反復処理
	while (enu->MoveNext()) {
		//テーブルを取得
		lo = (Excel::ListObject^)enu->Current;
		//開始位置は0ではなく、1から開始する
		testRange = (Excel::Range^)lo->Range->Cells[labelNum, beforeData];
		//テーブルに列を追加
		lo->ListColumns->Add(2);
		lo->ListColumns->Add(3);
		
		//ヘッダーを除き、最初の行から開始
		for (int i = 2; i < lo->ListRows->Count + 2; i++) {
			//加工対象データを取得
			testRange = (Excel::Range^)lo->Range->Cells[i, beforeData];
			String^ tmpData = testRange->Text->ToString();
			//4字目に区切り文字を挿入
			tmpData = tmpData->Insert(4, "#");
			//区切り文字の前後で分ける
			cli::array<String^>^ arr = tmpData->Split('#');
			//区切り文字の前半を2列目に代入
			lo->Range->Cells[i, 2] = arr[0];
			//区切り文字の後半を3列目に代入
			lo->Range->Cells[i, 3] = arr[1];

		}
		//ソート前にソート条件をクリアしておく
		lo->Sort->SortFields->Clear();
		//Rangeに2列目を指定
		testRange = (Excel::Range^)lo->ListColumns[2]->Range;
		try {
			//最優先である2列目のソート条件を指定
			lo->Sort->SortFields->Add2(
				testRange, Excel::XlSortOn::xlSortOnValues, Excel::XlSortOrder::xlAscending, Type::Missing, Excel::XlSortDataOption::xlSortNormal, Type::Missing);
		}catch(Exception^ e) {
			MessageBox::Show(e->ToString());
		}
		//rangeに3列目を指定
		testRange = (Excel::Range^)lo->ListColumns[3]->Range;
		//2番目に優先する3列目のソート条件を指定
		lo->Sort->SortFields->Add2(
			testRange, Excel::XlSortOn::xlSortOnValues, Excel::XlSortOrder::xlAscending, Type::Missing, Excel::XlSortDataOption::xlSortNormal, Type::Missing);
		//ソート実行のためのプロパティを設定
		lo->Sort->Header = XlYesNoGuess::xlYes;
		lo->Sort->MatchCase = false;
		lo->Sort->Orientation = XlSortOrientation::xlSortColumns;
		lo->Sort->SortMethod = XlSortMethod::xlPinYin;
		lo->Sort->Apply();

		break;
	}
	lo->ListColumns[3]->Delete();
	lo->ListColumns[2]->Delete();


	workbook->Save();

	//Excelのプロセスを閉じる処理
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
	//workbook->Close(Type::Missing, Type::Missing, Type::Missing);
	if (workbook != nullptr)
	{
		System::Runtime::InteropServices::Marshal::ReleaseComObject(workbook);
		workbook = nullptr;
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

	return System::Void();
}

System::Void ExcelClass::MyForm::createTable(int index, PowerPoint::Slide^ slide1, int height,int tableNum)
{
	
	PowerPoint::Shape^ containTable = slide1->Shapes->AddTable(10, 4, (350*(tableNum-1)), startY_value, 300, 300);
	PowerPoint::Table^ table = containTable->Table;
	
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

			//パワーポイント内のセルのテキストを中央寄せにする
			table->Rows[j]->Cells[i]->Shape->TextFrame->HorizontalAnchor = MsoHorizontalAnchor::msoAnchorCenter;
			//なお、数値や文字列などの書式の設定はエクセル機能での対応なのでパワーポイントでは指定しようがない
			if (i == 1 && j == 1) {
				//改行の区切り文字のテスト
				table->Rows[j]->Cells[i]->Shape->TextFrame->TextRange->Text = "100,200,300,400" + "\r\n" + "50,60,70,80";
				table->Rows[j]->Cells[i]->Shape->TextFrame->TextRange->Font->Size = FontSize;
				//フォントサイズの10倍に列の幅を設定
				table->Columns[i]->Width = FontSize * 10;
				//totalHeight += table->Rows[j]->Cells[i]->Shape->Height;
			}
			if (i == 2 && j == 2) {
				table->Rows[j]->Cells[i]->Shape->TextFrame->TextRange->Text = "sampleTest";
				table->Rows[j]->Cells[i]->Shape->TextFrame->TextRange->Font->Size = FontSize;
				//フォントサイズの10倍に列の幅を設定
				table->Columns[i]->Width = FontSize * 10;
				totalHeight += table->Rows[j]->Cells[i]->Shape->Height;
			}
			//セルの高さの更新
			if (tmpHeight < table->Rows[j]->Cells[i]->Shape->Height) {
				tmpHeight = table->Rows[j]->Cells[i]->Shape->Height;
			}
		}
	}
	return System::Void();
}
