# -*- coding: utf-8 -*-

import openpyxl as opx
import os
import sys
import json  
import types
import argparse


reload(sys)
sys.setdefaultencoding('utf-8')

specialField_list = ["one-off"]
jsonString_list = ["{", "}", "[", "]"]

def getExcelSheet(excelPath):
	workBook = opx.load_workbook(excelPath)
	return workBook

def isRightFirstColValue(value_str,type_client_or_server):
	if len(value_str) == 2:
		if value_str[1] == type_client_or_server:
			result = True
		elif value_str[1] == "-":
			result = True
		else:
			result = False
	elif len(value_str) == 1:
		result = True
	else:
		result = False
	return result

def isLegalValue(value_str,type_client_or_server):
	if value_str != "-":
		if type_client_or_server:
			isFormate = isRightFirstColValue(value_str,type_client_or_server)
			if isFormate:
				return True
		else:
			return True

def getDicOfFisrtCol(col_1st, type_client_or_server=None):
	dic = {}
	for k,v in enumerate(col_1st):
		if v:
			value_str = str(v)
			isLegal = isLegalValue(value_str,type_client_or_server)
			if isLegal:
				dic[k] = value_str[0]
	return dic

def isValueInRepeat(strValue,dic):
    if strValue in dic.values():
        isRepeat = True
    else:
        isRepeat = False
    return isRepeat

def addDataTo_dic(col2,dic):
    for k,v in enumerate(col2):
        strValue = str(v)
        if v:
            if isValueInRepeat(strValue,dic):
                raise Exception("=== Repeat Field! Sheet %s Row %s Col %s ==="%(str(sname), 2, k+1))
            else:
                dic[k] = strValue

def col2AppendDic(col_2nd,dic,index_list,sname):
	for k,v in enumerate(col_2nd):
		if v:
			strValue = str(v)
			if strValue in dic.values():
				raise Exception("=== Repeat Field! Sheet %s Row %s Col %s %s ==="%(str(sname), 2, k+1,strValue))
			else:
				dic[k] = strValue
	for k1,v1 in dic.items():
		specialFieldHandle(k1,v1,index_list)
	if len(index_list) > 0:
		for i in index_list:
			del dic[i]

def getDicOfSecondCol(col_2nd, type_lang=None,sname=None,temp=0):
	dic = {}
	temp_dic = {}
	formated_list=[]
	index_list = []
	if type_lang:
		addDataTo_dic(col_2nd,temp_dic)
		for v1 in temp_dic.values():
			addFormatedList(formated_list,v1)
		if len(formated_list) == 0:
			col2AppendDic(col_2nd,dic,index_list,sname)
		else:
			temp = getFormatedValueCount(temp_dic,type_lang)
			if temp == 0:
				col2AppendDic(col_2nd,dic,index_list,sname)
			else:
				for v2 in formated_list:
					for k3,v3 in temp_dic.items():
						if v2 == v3:
							index_list.append(k3)
						elif v2 in v3:
							formated1 = "-" + type_lang
							if formated1 not in v3:
								index_list.append(k3)
				delRepeatValue(index_list,temp_dic)
				delLinkSymbal(temp_dic)
				dic = temp_dic
	else:
		col2AppendDic(col_2nd,dic,index_list,sname)
	return dic

def specialFieldHandle(k,v,index):
	if v in specialField_list:
		pass
	elif "-" in v:
		index.append(k)

def addFormatedList(formated_list,v1):
    if "-" in v1:
        if v1.split("-")[0] not in formated_list:
            formated_list.append(v1.split("-")[0])
            
def getFormatedValueCount(dic,type2):
    formatedValueCount = 0
    for v in dic.values():
        formated1 = "-" + type2
        if formated1 in v:
        	formatedValueCount += 1
    return formatedValueCount 

def delRepeatValue(index,dic):
    for i in index:
        del dic[i]

def delLinkSymbal(dic):
    for k,v in dic.items():
        dic[k] = v.split("-")[0]

def indexSelect(col1,col2,t=[]):
    index = [i for i in col1.keys() if i in col2.keys()]
    col1Value = { i:col1[i] for i in index }
    col2Value = { i:col2[i] for i in index }
    t = [col1Value,col2Value]
    return (t, index)

def returnTypeInt(v):
	try:
		result = int(v)
	except:
		result = "\"" + v + "\""
	return result

def returnTypeJson(v):
	try:
		jsonload_value = json.loads(v)
		result = dic_to_lua_str(jsonload_value)
	except:
		result = "\"" + str(v) + "\""
	return result

def isStringInJsonFormat(str_value):
	for i in jsonString_list:
		if i in str_value:
			return False
	return True

def returnTypeMix(v):
	try:
		jsonload_value = json.loads(str(v))
		result = returnTypeJson(v)
	except Exception as e:
		temp = str(v)
		print temp
		isValueNum = temp.isdigit()
		if isValueNum:
			result = int(v)
		else:
			result = temp
	return result

def convertValue(type,v):
	if v is None:
		return "\"\""
	else:
		if type == "i":
			return returnTypeInt(v)
		elif type == "f":
			return float(v)
		elif type == "j":
			return returnTypeJson(v)
		elif type == "s":
			return stringTypeConvert(v)
		elif type == "m":
			return  returnTypeMix(v)

def stringTypeConvert(v):
	value = str(v)
	str_value = "\""
	for i in value:
		if i == '"' or i == "\\":
			temp = "\\%s"%i
		elif i == "\n":
			temp = "\\n"
		else:
			temp = i
		str_value += temp
	str_value += "\""
	return str_value

def space_str(layer):  
    lua_str = ""  
    for i in range(0,layer):  
        lua_str += '\t'  
    return lua_str  
  
def dic_to_lua_str(data,layer=0):  
    d_type = type(data)  
    if  d_type is types.StringTypes or d_type is str or d_type is types.UnicodeType:  
    	data = stringTypeConvert(data)
        return  data 
    elif d_type is types.BooleanType:  
        if data:  
            return 'true'  
        else:  
            return 'false'  
    elif d_type is types.IntType or d_type is types.LongType or d_type is types.FloatType:  
        return str(data)  
    elif d_type is types.ListType:  
        lua_str = "{"  
        for i in range(0,len(data)):  
            lua_str += dic_to_lua_str(data[i],layer+1)  
            if i < len(data) - 1:  
                lua_str += ','  
        lua_str +=  '}' 
        return lua_str  
    elif d_type is types.DictType:   
        lua_str = "{"  
        data_len = len(data)  
        data_count = 0  
        for k,v in data.items():  
            data_count += 1   
            str_key = str(k)
            if str_key.isdigit():  
                lua_str += '[' + k + ']'  
            else:  
                lua_str += "[\"%s\"]"%k   
            lua_str += '='    
            lua_str += dic_to_lua_str(v,layer +1)  
            if data_count < data_len:  
                lua_str += ','   
        lua_str += '}'  
        return lua_str  
    elif data is None:
    	return "nil" 

def convertToLua(type_list, title_list,index_list,valueTuple):
	string_format = "{"
	for k,v in enumerate(valueTuple):
		if k in index_list:
			_type = type_list[k]
			convert_value = convertValue(_type,v)
			temp = "[\"%s\"]=%s,"%(title_list[k],convert_value)
			string_format += temp
	string_format += "}"
	return string_format

def getFileNameWithoutSuffix(src):
	temp = src.split("/")
	fileName = temp[len(temp)-1].replace(".xlsx",".lua")
	return fileName

def main(src,type1,type2,path):
	fileName = src
	optpath = path + "/" + getFileNameWithoutSuffix(src) 
	excelSheet = getExcelSheet(fileName)
	excelValues = excelSheet.get_active_sheet().values
	with open(optpath,"wb") as f:
		f.write("return {\n")
		for k,v in enumerate(excelValues):
			if k == 1:
				firstCol = getDicOfFisrtCol(v,type_client_or_server=type1)
			if k == 2:
				secondCol = getDicOfSecondCol(v,type_lang=type2)
				col1_and_col2 = indexSelect(firstCol,secondCol)[0] 
				value_index = indexSelect(firstCol,secondCol)[1] 
			if k >=3 :
				table = convertToLua(col1_and_col2[0],col1_and_col2[1],value_index,v)
				f.write(table + ",\n")
		f.write("}")

def saveLua(src,path,filename,type1=None,type2=None):
	filepath = src
	luaFile = path + "/" + filename+ ".lua"
	excelFile = src + "/" + filename + ".xlsx"
	excelSheet = getExcelSheet(excelFile)
	excelValues = excelSheet.get_active_sheet().values
	with open(luaFile,"wb") as f:
		f.write("return {\n")
		for k,v in enumerate(excelValues):
			if k == 1:
				firstCol = getDicOfFisrtCol(v,type_client_or_server=type1)
			if k == 2:
				secondCol = getDicOfSecondCol(v,type_lang=type2)
				col1_and_col2 = indexSelect(firstCol,secondCol)[0] 
				value_index = indexSelect(firstCol,secondCol)[1] 
			if k >=3 :
				table = convertToLua(col1_and_col2[0],col1_and_col2[1],value_index,v)
				f.write(table + ",\n")
		f.write("}")

def readAllExcelAndSaveLua(src,path,type1,type2):
	for parent, dirs, files in os.walk(src):
		for fn in files:
			if "~$" in fn:   # temporary excel can not read in 
				continue
			else:
				(fileName, suffix) = os.path.splitext(fn)
				saveLua(src,path,fileName,type1,type2)

def getParserArgument():
	parser = argparse.ArgumentParser(description='convert excel to lua')
	parser.add_argument("excelPath", metavar="src",help="excel path")
	parser.add_argument("-o",metavar="output",dest="optPath",help="output path")
	parser.add_argument("-l",metavar="lang",dest="language",help="field language")
	parser.add_argument("-c",dest="forClient",action='store_const',const=True,default=False,help="dump for client")
	parser.add_argument("-s",dest="forServer",action='store_const',const=True,default=False,help="dump for server")
	args = parser.parse_args()
	return args

if __name__ == "__main__":
    args = getParserArgument()
    src = args.excelPath
    if args.optPath:
        path = str(args.optPath)
    if args.forClient:
        type1 = 'c'
    elif args.forServer:
        type1 = 's'
    else:
        type1 = None
    if args.language:
        type2 = str(args.language)
    else:
        type2 = None

    if ".xlsx" in src:
    	main(src,type1,type2,path)
    else:
    	readAllExcelAndSaveLua(src,path,type1,type2)
    print "All Done"



