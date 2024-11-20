import os
import re

class Util:

    # 清空文件夹
    @staticmethod
    def clear_folder(dir_path):
        for root, dirs, files in os.walk(dir_path, topdown=False):
            # 第一步：删除文件
            for name in files:
                os.remove(os.path.join(root, name))  # 删除文件
            # 第二步：删除空文件夹
            for name in dirs:
                os.rmdir(os.path.join(root, name)) # 删除一个空目录

    # 字典安全取值
    @staticmethod
    def safe_value(dictionary, key):
        if isinstance(dictionary, dict) and key in dictionary:
            return dictionary[key]
        else:
            return None 
        
    # 根据text创建一个在map中不重复的key
    @staticmethod
    def create_not_repeat_key_in_map(text, key_map, key_word_count):
        # 取前key_word_count个单词
        words = re.split(r"[,.:：“”<>‘ + ’\d%…?.!***“\\_/\-\&\n{}()/'\"\s]+", text)
        words = list(filter(None, words))
        words_len = len(words)
        key_word = words
        if words_len > key_word_count:
            key_word = words[:key_word_count]
        key = '_'.join(key_word)    
        _key = key.lower()

        if key_map is None:
            return _key
        
        filter_key_count = sum(1 for k,v in key_map.items() if v == _key)
        # key已经存在map中，则网后多取一个单词
        if filter_key_count > 0: 
            # 生成key的单词比原句少
            if key_word_count < words_len: 
                new_key = None
                while new_key is None:
                    key_word_count += 1
                    new_key = Util.create_not_repeat_key_in_map(text, key_map, key_word_count)
                return new_key
            # 生成key的单词跟原句一样，则添加下标区分
            else: 
                return _key + '_' + f'{filter_key_count + 1}'
                
        else: # key不存在map中，则返回新生成的key
            return _key
        
    