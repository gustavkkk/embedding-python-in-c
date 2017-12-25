// MoreComplex.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "Python.h"

void Extract();

int main()
{
	Extract();
    return 0;
}

void Extract()
{
	PyObject *pName, *pModule = NULL, *pDict, *pFunc = NULL;
	PyObject *pArgs, *pValue;

	Py_Initialize();
	//PySys_SetPath(_T("libs"));
	pName = PyUnicode_DecodeFSDefault("extract");
	/* Error checking of pName left out */
	pModule = PyImport_Import(pName);
	Py_DECREF(pName);

	if (pModule != NULL) {
		//const char* func_name2 = 'multiply';
		pFunc = PyObject_GetAttrString(pModule, "extract");
		/* pFunc is a new reference */

		if (pFunc && PyCallable_Check(pFunc)) {
			//set arg
			pArgs = PyTuple_New(1);
			PyTuple_SetItem(pArgs, 0, Py_BuildValue("s", "simohua.docx"));
			//get result
			pValue = PyObject_CallObject(pFunc, pArgs);
			Py_DECREF(pArgs);
			Py_DECREF(pModule);
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