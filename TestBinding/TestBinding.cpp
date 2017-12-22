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

//调用输出"Hello World"函数  
void HelloWorld()
{
	Py_Initialize();//使用python之前，要调用Py_Initialize();这个函数进行初始化 
	PyObject * pModule = NULL;//声明变量  
	PyObject * pFunc = NULL;//声明变量
#ifdef CONFIDENT
	pModule = PyImport_ImportModule("test");//这里是要调用的Python文件名 
#else
	pModule = PyImport_Import(PyUnicode_DecodeFSDefault("script"));
#endif
	if (!pModule) {
		printf("Cant open python file!\n");
		return;
	}
	pFunc = PyObject_GetAttrString(pModule, "HelloWorld"); //这里是要调用的函数名  
	if (!pFunc) {
		cout << "Get function failed" << endl;
		return;
	}
	PyEval_CallObject(pFunc, NULL); //调用函数,NULL表示参数为空  
	Py_Finalize();//调用Py_Finalize,这个和Py_Initialize相对应的.  
}

//调用Add函数,传两个int型参数  
void Add()
{
	Py_Initialize();

	PyObject * pModule = NULL;
	PyObject * pFunc = NULL;
	pModule = PyImport_ImportModule("script");//test:Python文件名  
	pFunc = PyObject_GetAttrString(pModule, "add");//Add:Python文件中的函数名    //创建参数  
	PyObject *pArgs = PyTuple_New(2); //函数调用的参数传递均是以元组的形式打包的,2表示参数个数  
	PyTuple_SetItem(pArgs, 0, Py_BuildValue("i", 5));//0---序号i表示创建int型变量  
	PyTuple_SetItem(pArgs, 1, Py_BuildValue("i", 7));//1---序号//返回值  
	PyObject *pReturn = NULL;
	pReturn = PyEval_CallObject(pFunc, pArgs);//调用函数  
											  //将返回值转换为int类型  
	int result;
	PyArg_Parse(pReturn, "i", &result);//i表示转换成int型变量  
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

//参数传递的类型为字典  
void TestTransferDict()
{
	Py_Initialize();

	PyObject * pModule = NULL;
	PyObject * pFunc = NULL;
	pModule = PyImport_ImportModule("script");//test:Python文件名  
	pFunc = PyObject_GetAttrString(pModule, "TestDict"); //Add:Python文件中的函数名  
														 //创建参数:  
	PyObject *pArgs = PyTuple_New(1);
	PyObject *pDict = PyDict_New(); //创建字典类型变量  
	PyDict_SetItemString(pDict, "Name", Py_BuildValue("s", "WangYao")); //往字典类型变量中填充数据  
	PyDict_SetItemString(pDict, "Age", Py_BuildValue("i", 25)); //往字典类型变量中填充数据  
	PyTuple_SetItem(pArgs, 0, pDict);//0---序号将字典类型变量添加到参数元组中  
									 //返回值  
	PyObject *pReturn = NULL;
	pReturn = PyEval_CallObject(pFunc, pArgs);//调用函数  
											  //处理返回值:  
	int size = PyDict_Size(pReturn);
	cout << "返回字典的大小为: " << size << endl;
	PyObject *pNewAge = PyDict_GetItemString(pReturn, "Age");
	int newAge;
	PyArg_Parse(pNewAge, "i", &newAge);
	cout << "True Age: " << newAge << endl;

	Py_Finalize();
}

//测试类  
void TestClass()
{
	Py_Initialize();

	PyObject * pModule = NULL;
	PyObject * pFunc = NULL;
	pModule = PyImport_ImportModule("script");//test:Python文件名  
	pFunc = PyObject_GetAttrString(pModule, "TestDict"); //Add:Python文件中的函数名  
														 //获取Person类  
	PyObject *pClassPerson = PyObject_GetAttrString(pModule, "Person");
	//创建Person类的实例  
	//PyObject *pInstancePerson = PyInstance_New(pClassPerson, NULL, NULL);
	//调用方法  
	//PyObject_CallMethod(pInstancePerson, "greet", "s", "Hello Kitty"); //s表示传递的是字符串,值为"Hello Kitty"  

	Py_Finalize();
}
