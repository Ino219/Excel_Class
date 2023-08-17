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
	ProgressBar^ pBar1 = progressBar1;
	int Min = 1;
	int Max = 10;
	int Value = 1;

	//�ŏ��A�ő�l�̐ݒ�
	pBar1->Minimum = Min;
	pBar1->Maximum = Max;
	//�����l�̐ݒ�
	pBar1->Value = Min;
	//���̐ݒ�
	pBar1->Step = Value;

	//�v���O���X�o�[��i�s������
	for (int i = Min; i < Max; i++) {
		System::Threading::Thread::Sleep(1000);
		pBar1->PerformStep();
	}
	return System::Void();
}
