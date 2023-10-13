#include "MyForm.h"

using namespace ProgressBarSample;

[STAThreadAttribute]

int main() {
	Application::Run(gcnew MyForm());
	return 0;
}

System::Void ProgressBarSample::MyForm::MyForm_Load_(System::Object ^ sender, System::EventArgs ^ e)
{
	
	return System::Void();
}

System::Void ProgressBarSample::MyForm::button1_Click(System::Object ^ sender, System::EventArgs ^ e)
{
	logStr += "start" + "\r\n";
	ProgressBar^ pBar1 = progressBar1;
	int Min = 1;
	int Max = 10;
	int Value = 1;

	//最小、最大値の設定
	pBar1->Minimum = Min;
	pBar1->Maximum = Max;
	//初期値の設定
	pBar1->Value = Min;
	//幅の設定
	pBar1->Step = Value;

	//プログレスバーを進行させる
	for (int i = Min; i < Max; i++) {
		System::Threading::Thread::Sleep(1000);
		logStr += "ステータスバー更新"+"\r\n";
		pBar1->PerformStep();
	}
	return System::Void();
}
