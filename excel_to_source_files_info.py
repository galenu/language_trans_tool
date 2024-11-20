from collections import OrderedDict

# 设备类型
class Platform:
    
    ios = 'ios' # ios
    
    android = 'android' # android
    
    web = 'web' # web

# 项目类型 安卓国际化文件名每个项目不一致,需要根据项目配置
class Project:

    # partybox 或 partylight
    partybox = 'partybox'  
    
    # 新的jblone_3_in_1
    jblone = 'jblone' 

    # party_box和party_light的所有语言
    @staticmethod
    def _partybox_languages():
        return [
            ExcelToSourceFilesInfo('en', '(English)', 'en.lproj', 'values', 'en.js'), # 其他语言没找到，默认会从第一个语言查找
            ExcelToSourceFilesInfo('fr', '(French)', 'fr.lproj', 'values-fr', 'fr.js'),
            ExcelToSourceFilesInfo('da', '(Danish)', 'da.lproj', 'values-da', 'da.js'),
            ExcelToSourceFilesInfo('de', '(German)', 'de.lproj', 'values-de', 'de.js'),
            ExcelToSourceFilesInfo('es', '(Spanish)', 'es.lproj', 'values-es', 'es.js'),
            ExcelToSourceFilesInfo('es-MX', '(Spanish_Mexico)', 'es-MX.lproj', 'values-es-rMX', 'es-MX.js'),
            ExcelToSourceFilesInfo('fi', '(Finnish)', 'fi.lproj', 'values-fi', 'fi.js'),
            ExcelToSourceFilesInfo('it', '(Italian)', 'it.lproj', 'values-it', 'it.js'),
            ExcelToSourceFilesInfo('nl', '(Dutch)', 'nl.lproj', 'values-nl', 'nl.js'),
            ExcelToSourceFilesInfo('no', '(Norwegian)', 'nb.lproj', 'values-no', 'nb.js'),
            ExcelToSourceFilesInfo('pl', '(Polish)', 'pl.lproj', 'values-pl', 'pl.js'),
            ExcelToSourceFilesInfo('pt-BR', '(Portuguese_Brazil)', 'pt-BR.lproj', 'values-pt-rBR', 'pt-BR.js'),
            ExcelToSourceFilesInfo('pt-PT', '(Portuguese_portugal)', 'pt-PT.lproj', 'values-pt-rPT', 'pt-PT.js'),
            ExcelToSourceFilesInfo('sk', '(Slovak)', 'sk.lproj', 'values-sk', 'ar.sk'),
            ExcelToSourceFilesInfo('sv', '(Swedish)', 'sv.lproj', 'values-sv', 'sv.js'),
            ExcelToSourceFilesInfo('ru', '(Russian)', 'ru.lproj', 'values-ru', 'ru.js'),
            ExcelToSourceFilesInfo('zh-Hans', '(Simplified Chinese)', 'zh-Hans.lproj', 'values-zh-rCN', 'zh-Hans.js'),
            ExcelToSourceFilesInfo('zh-Hant', '(Traditional Chinese)', 'zh-Hant.lproj', 'values-zh-rTW', 'zh-Hant.js'),
            ExcelToSourceFilesInfo('id', '(Indonesian)', 'id.lproj', 'values-in-rID', 'id.js'),
            ExcelToSourceFilesInfo('jp', '(Japanese)', 'ja.lproj', 'values-ja', 'ja.js'),
            ExcelToSourceFilesInfo('ko', '(Korean)', 'ko.lproj', 'values-ko', 'ko.js'),
            ExcelToSourceFilesInfo('ar', '(Arabic)', 'ar.lproj', 'values-ar', 'ar.js'),
            ExcelToSourceFilesInfo('he', '(Hebrew)', 'he.lproj', 'values-iw', 'he.js')
        ]
    
    # 新的jblone_3_in_1的所有语言
    @staticmethod
    def _jblone_languages():
        return [
            ExcelToSourceFilesInfo('en', '(English)', 'en.lproj', 'values', 'en.js'), # 其他语言没找到，默认会从第一个语言查找
            ExcelToSourceFilesInfo('fr', '(French)', 'fr.lproj', 'values-fr-rFR', 'fr.js'),
            ExcelToSourceFilesInfo('da', '(Danish)', 'da.lproj', 'values-da-rDK', 'da.js'),
            ExcelToSourceFilesInfo('de', '(German)', 'de.lproj', 'values-de-rDE', 'de.js'),
            ExcelToSourceFilesInfo('es', '(Spanish)', 'es.lproj', 'values-es-rES', 'es.js'),
            ExcelToSourceFilesInfo('es-MX', '(Spanish_Mexico)', 'es-MX.lproj', 'values-es-rMX', 'es-MX.js'),
            ExcelToSourceFilesInfo('fi', '(Finnish)', 'fi.lproj', 'values-fi-rFI', 'fi.js'),
            ExcelToSourceFilesInfo('it', '(Italian)', 'it.lproj', 'values-it-rIT', 'it.js'),
            ExcelToSourceFilesInfo('nl', '(Dutch)', 'nl.lproj', 'values-nl-rNL', 'nl.js'),
            ExcelToSourceFilesInfo('no', '(Norwegian)', 'nb.lproj', 'values-nb-rNO', 'nb.js'),
            ExcelToSourceFilesInfo('pl', '(Polish)', 'pl.lproj', 'values-pl-rPL', 'pl.js'),
            ExcelToSourceFilesInfo('pt-BR', '(Portuguese_Brazil)', 'pt-BR.lproj', 'values-pt-rBR', 'pt-BR.js'),
            ExcelToSourceFilesInfo('pt-PT', '(Portuguese_portugal)', 'pt-PT.lproj', 'values-pt-rPT', 'pt-PT.js'),
            ExcelToSourceFilesInfo('sk', '(Slovak)', 'sk.lproj', 'values-sk-rSK', 'sk.js'),
            ExcelToSourceFilesInfo('sv', '(Swedish)', 'sv.lproj', 'values-sv-rSE', 'sv.js'),
            ExcelToSourceFilesInfo('ru', '(Russian)', 'ru.lproj', 'values-ru-rRU', 'ru.js'),
            ExcelToSourceFilesInfo('zh-Hans', '(Simplified Chinese)', 'zh-Hans.lproj', 'values-zh-rCN', 'zh-Hans.js'),
            ExcelToSourceFilesInfo('zh-Hant', '(Traditional Chinese)', 'zh-Hant.lproj', 'values-zh-rTW', 'zh-Hant.js'),
            ExcelToSourceFilesInfo('id', '(Indonesian)', 'id.lproj', 'values-in-rID', 'id.js'),
            ExcelToSourceFilesInfo('jp', '(Japanese)', 'ja.lproj', 'values-ja-rJP', 'ja.js'),
            ExcelToSourceFilesInfo('ko', '(Korean)', 'ko.lproj', 'values-ko-rKR', 'ko.js'),
            ExcelToSourceFilesInfo('ar', '(Arabic)', 'ar.lproj', 'values-ar-rAE', 'ar.js'),
            ExcelToSourceFilesInfo('he', '(Hebrew)', 'he.lproj', 'values-iw-rIL', 'he.js')     
        ]
    
    # 所有语言
    @staticmethod
    def allLanguages(project): 
        if project == Project.jblone:
            return Project._jblone_languages()
        if project == Project.partybox:
            return Project._partybox_languages()
        
    # 转义符 excel中转到源码需要加转义符的 字符
    @staticmethod
    def getEscapedCharacters(): 
        return ['"', '\'', '‘', '’', '“', '”']
    
# excel与转换源文件的映射表
class ExcelToSourceFilesInfo:

    # 构造函数（或称为初始化方法）
    def __init__(self, excel_lang_id, desc, ios_lproj_file_name, android_values_file_name, web_js_file_name):
        self.excel_lang_id = excel_lang_id # excel上语言对应的id(取第一行每列:之前的) en:英语(English)
        self.desc = desc # excel上语言对应的描述(取第一行每列:之后的)
        self.ios_lproj_file_name = ios_lproj_file_name # ios国际化源文件夹名(en.lproj)
        self.android_values_file_name = android_values_file_name # android国际化源文件夹名(values-zh-rCN)
        self.web_js_file_name = web_js_file_name # web国际化源文件名(en.js)
