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
		//インスタンスの設定
		pBar1 = gcnew ProgressBar();
	}
	static System::Void setValue(int Min, int Max, int Value) {
		//最小、最大値の設定
		pBar1->Minimum = Min;
		pBar1->Maximum = Max;
		//初期値の設定
		pBar1->Value = Min;
		//幅の設定
		pBar1->Step = Value;
	}
	static System::Void Perfomed() {
		//プログレスバーを進行させる
		pBar1->PerformStep();
	}
	
};

