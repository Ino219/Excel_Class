#pragma once

using namespace System;
using namespace System::Windows::Forms;


ref class Class
{
	
	Class() {
		
	}

	public:

		static int main() {
			return 0;
	}

	static ProgressBar^ pBar1;

	static System::Void InstanceProgressBar() {
		//�C���X�^���X�̐ݒ�
		pBar1 = gcnew ProgressBar();
	}
	static System::Void setValue(int Min, int Max, int Value) {
		//�ŏ��A�ő�l�̐ݒ�
		pBar1->Minimum = Min;
		pBar1->Maximum = Max;
		//�����l�̐ݒ�
		pBar1->Value = Min;
		//���̐ݒ�
		pBar1->Step = Value;
	}
	static System::Void Perfomed() {
		//�v���O���X�o�[��i�s������
		pBar1->PerformStep();
	}
	
};

