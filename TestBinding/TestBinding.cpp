// TestBinding.cpp : Defines the entry point for the console application.
// http://blog.csdn.net/c_cyoxi/article/details/23978007
// Don't name python file 'test.py'. It doesn't work.
// https://stackoverflow.com/questions/24313681/pyobject-getattrstring-c-function-returning-null-unable-to-call-python-functi
#include "stdafx.h"
#include <iostream>  
#include <Python.h>  

using namespace std;

void Simplest();
void HelloWorld();
void Add();
void Multiply();
void TestTransferDict();
void TestClass();

int main()
{
	cout << "Starting Test..." << endl;
	cout << "Simplest()-------------" << endl;
	Simplest();
	cout << "HelloWorld()-------------" << endl;
	HelloWorld();
	cout << "Add()--------------------" << endl;
	Add();
	cout << "Multiply()--------------------" << endl;
	Multiply();
	cout << "TestDict-----------------" << endl;
	TestTransferDict();
	cout << "TestClass----------------" << endl;
	TestClass();

	system("pause");
	return 0;
}

/*
wchar_t* GetFileName(char* argv[])
{
	wchar_t progname[FILENAME_MAX + 1];
	mbstowcs(progname, argv[0], strlen(argv[0]) + 1);
	return progname;
}
*/
void Simplest()
{
	wchar_t *title = _T("simple");
	Py_SetProgramName(title);//argv[0]);
	Py_Initialize();
	char *string = "from time import time,ctime\n""print('Today is', ctime(time()))\n";
	PyRun_SimpleString(string);
	Py_Finalize();
	return;
}

//�������"Hello World"����  
void HelloWorld()
{
	Py_Initialize();//ʹ��python֮ǰ��Ҫ����Py_Initialize();����������г�ʼ�� 
	PyObject * pModule = NULL;//��������  
	PyObject * pFunc = NULL;//��������
#ifdef CONFIDENT
	pModule = PyImport_ImportModule("test");//������Ҫ���õ�Python�ļ��� 
#else
	pModule = PyImport_Import(PyUnicode_DecodeFSDefault("script"));
#endif
	if (!pModule) {
		printf("Cant open python file!\n");
		return;
	}
	pFunc = PyObject_GetAttrString(pModule, "HelloWorld"); //������Ҫ���õĺ�����  
	if (!pFunc) {
		cout << "Get function failed" << endl;
		return;
	}
	PyEval_CallObject(pFunc, NULL); //���ú���,NULL��ʾ����Ϊ��  
	Py_Finalize();//����Py_Finalize,�����Py_Initialize���Ӧ��.  
}

//����Add����,������int�Ͳ���  
void Add()
{
	Py_Initialize();

	PyObject * pModule = NULL;
	PyObject * pFunc = NULL;
	pModule = PyImport_ImportModule("script");//test:Python�ļ���  
	pFunc = PyObject_GetAttrString(pModule, "add");//Add:Python�ļ��еĺ�����    //��������  
	PyObject *pArgs = PyTuple_New(2); //�������õĲ������ݾ�����Ԫ�����ʽ�����,2��ʾ��������  
	PyTuple_SetItem(pArgs, 0, Py_BuildValue("i", 5));//0---���i��ʾ����int�ͱ���  
	PyTuple_SetItem(pArgs, 1, Py_BuildValue("i", 7));//1---���//����ֵ  
	PyObject *pReturn = NULL;
	pReturn = PyEval_CallObject(pFunc, pArgs);//���ú���  
											  //������ֵת��Ϊint����  
	int result;
	PyArg_Parse(pReturn, "i", &result);//i��ʾת����int�ͱ���  
	cout << "5+7 = " << result << endl;

	Py_Finalize();
}

void Multiply()
{
	PyObject *pName, *pModule = NULL, *pDict, *pFunc = NULL;
	PyObject *pArgs, *pValue;

	Py_Initialize();
	pName = PyUnicode_DecodeFSDefault("script");
	/* Error checking of pName left out */
	pModule = PyImport_Import(pName);
	Py_DECREF(pName);

	if (pModule != NULL) {
		//const char* func_name2 = 'multiply';
		pFunc = PyObject_GetAttrString(pModule, "multiply");
		/* pFunc is a new reference */

		if (pFunc && PyCallable_Check(pFunc)) {
			pArgs = PyTuple_New(2);
			int value[2] = {3,4};
			for (int i = 0; i < 2; ++i) {
				pValue = PyLong_FromLong(value[i]);// atoi(argv[i + 3]));
				if (!pValue) {
					Py_DECREF(pArgs);
					Py_DECREF(pModule);
					fprintf(stderr, "Cannot convert argument\n");
					return;
				}
				/* pValue reference stolen here: */
				PyTuple_SetItem(pArgs, i, pValue);
			}
			pValue = PyObject_CallObject(pFunc, pArgs);
			Py_DECREF(pArgs);
			if (pValue != NULL) {
				printf("Result of call: %ld\n", PyLong_AsLong(pValue));
				Py_DECREF(pValue);
			}
			else {
				Py_DECREF(pFunc);
				Py_DECREF(pModule);
				PyErr_Print();
				fprintf(stderr, "Call failed\n");
				return;
			}
		}
		else {
			if (PyErr_Occurred())
				PyErr_Print();
			fprintf(stderr, "Cannot find function \"%s\"\n", "multiply");
		}
		Py_XDECREF(pFunc);
		Py_DECREF(pModule);
	}
	else {
		PyErr_Print();
		fprintf(stderr, "Failed to load \"%s\"\n", "test");
		return;
	}
	Py_Finalize();
	return;
}

//�������ݵ�����Ϊ�ֵ�  
void TestTransferDict()
{
	Py_Initialize();

	PyObject * pModule = NULL;
	PyObject * pFunc = NULL;
	pModule = PyImport_ImportModule("script");//test:Python�ļ���  
	pFunc = PyObject_GetAttrString(pModule, "TestDict"); //Add:Python�ļ��еĺ�����  
														 //��������:  
	PyObject *pArgs = PyTuple_New(1);
	PyObject *pDict = PyDict_New(); //�����ֵ����ͱ���  
	PyDict_SetItemString(pDict, "Name", Py_BuildValue("s", "WangYao")); //���ֵ����ͱ������������  
	PyDict_SetItemString(pDict, "Age", Py_BuildValue("i", 25)); //���ֵ����ͱ������������  
	PyTuple_SetItem(pArgs, 0, pDict);//0---��Ž��ֵ����ͱ�����ӵ�����Ԫ����  
									 //����ֵ  
	PyObject *pReturn = NULL;
	pReturn = PyEval_CallObject(pFunc, pArgs);//���ú���  
											  //������ֵ:  
	int size = PyDict_Size(pReturn);
	cout << "�����ֵ�Ĵ�СΪ: " << size << endl;
	PyObject *pNewAge = PyDict_GetItemString(pReturn, "Age");
	int newAge;
	PyArg_Parse(pNewAge, "i", &newAge);
	cout << "True Age: " << newAge << endl;

	Py_Finalize();
}

//������  
void TestClass()
{
	Py_Initialize();

	PyObject * pModule = NULL;
	PyObject * pFunc = NULL;
	pModule = PyImport_ImportModule("script");//test:Python�ļ���  
	pFunc = PyObject_GetAttrString(pModule, "TestDict"); //Add:Python�ļ��еĺ�����  
														 //��ȡPerson��  
	PyObject *pClassPerson = PyObject_GetAttrString(pModule, "Person");
	//����Person���ʵ��  
	//PyObject *pInstancePerson = PyInstance_New(pClassPerson, NULL, NULL);
	//���÷���  
	//PyObject_CallMethod(pInstancePerson, "greet", "s", "Hello Kitty"); //s��ʾ���ݵ����ַ���,ֵΪ"Hello Kitty"  

	Py_Finalize();
}
