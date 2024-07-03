# -*- coding: utf-8 -*-
"""
Created on Tue Jun 11 10:38:34 2024

@author: liuchangjun
"""
#%% 初始配置
import xml.etree.ElementTree as ET
import xlwings as xw
from pathlib import Path

def get_initials(sentence):
    words = sentence.split()
    initials = [word[0] for word in words]
    return ''.join(initials)

def check_name(pre:list, node:list, name_node:str, no_fix_name_node:list, sheet_name:str, starx_idx=0):
    '''
    检查node/name_node是否等于pre+no_fix_name_node

    Parameters
    ----------
    pre : list
        no_fix_name_node节点的前缀字符，前缀以列表存储.
    node : list
        父节点路径，列表形式[父节点的父节点，父节点].
    name_node : str
        包含前缀的节点.
    no_fix_name_node : list
        无前缀的节点（若最终节点tag与name_node相同，则直接访问父节点下列表的第二个节点；否则以完整路径访问）.
    sheet_name : str
        新建的子表名.
    starx_idx : int
        有些不能直接比较no_fix_name_node， 使用start_idx截取前缀后的内容
    Returns
    -------
    None.

    '''
    node_name_tag = get_initials(no_fix_name_node[0].replace('-', ' '))
    if len(node) == 2:
        SenderReceiverInterfaces = root.findall(f".//{AB_DIR}{node[0]}/{AB_DIR}{node[1]}")
        node_tag = get_initials(node[0].replace('-', ' '))+'/'+get_initials(node[1].replace('-', ' '))
    elif len(node) == 1:
        SenderReceiverInterfaces = root.findall(f".//{AB_DIR}{node[0]}")
        node_tag = get_initials(node[0].replace('-', ' '))
    else:
        return None
    
    SHORT_NAMES, NO_PREFIX_SHORT_NAMES, UUID = [], [], []   
    for SenderReceiverInterface in SenderReceiverInterfaces:
        # print(f".//{AB_DIR}{name_node}")
        try:
            SHORT_NAME = SenderReceiverInterface.find(f".//{AB_DIR}{name_node}").text
            try:
                if name_node == no_fix_name_node[1]:
                    NO_PREFIX_SHORT_NAME = SenderReceiverInterface.find(f".//{AB_DIR}{no_fix_name_node[0]}/{AB_DIR}{no_fix_name_node[1]}").text
                else:
                    NO_PREFIX_SHORT_NAME = SenderReceiverInterface.find(f".//{AB_DIR}{no_fix_name_node[1]}").text
                NO_PREFIX_SHORT_NAME = NO_PREFIX_SHORT_NAME.split('/')[-1]
                SHORT_NAME = SHORT_NAME.split('/')[-1]
                    
                if SHORT_NAME[starx_idx:] not in [i + NO_PREFIX_SHORT_NAME for i in pre]:
                    SHORT_NAMES.append(SHORT_NAME[starx_idx:])
                    NO_PREFIX_SHORT_NAMES.append(NO_PREFIX_SHORT_NAME)
                    UUID.append(SenderReceiverInterface.get('UUID'))
            except:
               SHORT_NAMES.append(SHORT_NAME)
               NO_PREFIX_SHORT_NAMES.append(f'Not defined {no_fix_name_node[1]}')
               UUID.append(SenderReceiverInterface.get('UUID'))
        except:
            SHORT_NAMES.append(SenderReceiverInterface.find(f".//{AB_DIR}SHORT-NAME").text)
            NO_PREFIX_SHORT_NAMES.append(f'Not defined {name_node}')
            UUID.append(SenderReceiverInterface.get('UUID'))
            
    try:
        sheet_ = wb.sheets.add(sheet_name, after=wb.sheets[-1]) 
        sheet_.range('A1').value = ['INDEX', f'{node_tag}/{name_node}', f'{node_tag}/{node_name_tag}/{no_fix_name_node[1]}', 'UUID']  
        sheet_.range('A1').expand(mode="right").font.bold = True
        sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), SHORT_NAMES, NO_PREFIX_SHORT_NAMES, UUID]
        sheet_.range('A1').expand(mode='table').columns.autofit()
    except Exception as err:
        print(err)
        
def check_name_IB(pre:list, node:list, name_node:str, no_fix_name_node:list, sheet_name:str, starx_idx=0):
    '''
    检查node/name_node是否等于pre+no_fix_name_node

    Parameters
    ----------
    pre : list
        no_fix_name_node节点的前缀字符，前缀以列表存储.
    node : list
        父节点路径，列表形式[父节点的父节点，父节点].
    name_node : str
        包含前缀的节点.
    no_fix_name_node : list
        无前缀的节点（若最终节点tag与name_node相同，则直接访问父节点下列表的第二个节点；否则以完整路径访问）.
    sheet_name : str
        新建的子表名.
    starx_idx : int
        有些不能直接比较no_fix_name_node， 使用start_idx截取前缀后的内容
    Returns
    -------
    None.

    '''
    node_name_tag = get_initials(no_fix_name_node[0].replace('-', ' '))
    if len(node) == 2:
        SenderReceiverInterfaces = root.findall(f".//{AB_DIR}{node[0]}/{AB_DIR}{node[1]}")
        node_tag = get_initials(node[0].replace('-', ' '))+'/'+get_initials(node[1].replace('-', ' '))
    elif len(node) == 1:
        SenderReceiverInterfaces = root.findall(f".//{AB_DIR}{node[0]}")
        node_tag = get_initials(node[0].replace('-', ' '))
    else:
        return None
    
    SHORT_NAMES, NO_PREFIX_SHORT_NAMES, UUID = [], [], []   
    for SenderReceiverInterface in SenderReceiverInterfaces:
        # print(f".//{AB_DIR}{name_node}")
        try:
            SHORT_NAME = SenderReceiverInterface.find(f".//{AB_DIR}{name_node}").text
            try:
                if name_node == no_fix_name_node[1]:
                    NO_PREFIX_SHORT_NAME = SenderReceiverInterface.find(f".//{AB_DIR}{no_fix_name_node[0]}/{AB_DIR}{no_fix_name_node[1]}").text
                else:
                    NO_PREFIX_SHORT_NAME = SenderReceiverInterface.find(f".//{AB_DIR}{no_fix_name_node[1]}").text
                NO_PREFIX_SHORT_NAME = NO_PREFIX_SHORT_NAME.split('/')[-1]
                SHORT_NAME = SHORT_NAME.split('/')[-1]
                    
                if SHORT_NAME not in [i + NO_PREFIX_SHORT_NAME[starx_idx:] for i in pre]:
                    SHORT_NAMES.append(SHORT_NAME)
                    NO_PREFIX_SHORT_NAMES.append(NO_PREFIX_SHORT_NAME)
                    UUID.append(SenderReceiverInterface.get('UUID'))
            except:
               SHORT_NAMES.append(SHORT_NAME)
               NO_PREFIX_SHORT_NAMES.append(f'Not defined {no_fix_name_node[1]}')
               UUID.append(SenderReceiverInterface.get('UUID'))
        except:
            SHORT_NAMES.append(SenderReceiverInterface.find(f".//{AB_DIR}SHORT-NAME").text)
            NO_PREFIX_SHORT_NAMES.append(f'Not defined {name_node}')
            UUID.append(SenderReceiverInterface.get('UUID'))
            
    try:
        sheet_ = wb.sheets.add(sheet_name, after=wb.sheets[-1]) 
        sheet_.range('A1').value = ['INDEX', f'{node_tag}/{name_node}', f'{node_tag}/{node_name_tag}/{no_fix_name_node[1]}', 'UUID']  
        sheet_.range('A1').expand(mode="right").font.bold = True
        sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), SHORT_NAMES, NO_PREFIX_SHORT_NAMES, UUID]
        sheet_.range('A1').expand(mode='table').columns.autofit()
    except Exception as err:
        print(err)
        
def SR_port_check_name(pre:list, node:list, name_node:str, no_fix_name_node:list, child:str, attrib:str, sheet_name:str, starx_idx=0):
    node_tag = get_initials(node[0].replace('-', ' '))+'/'+get_initials(node[1].replace('-', ' '))
    node_name_tag = get_initials(no_fix_name_node[0].replace('-', ' '))
    
    SHORT_NAMES, NO_PREFIX_SHORT_NAMES, UUID = [], [], []   
    SenderReceiverInterfaces = root.findall(f".//{AB_DIR}{node[0]}/{AB_DIR}{node[1]}")
    for SenderReceiverInterface in SenderReceiverInterfaces:
        ATTRIB = SenderReceiverInterface.find(f'.//{AB_DIR}{child}')
        if ATTRIB.get('DEST') == attrib:
            try:
                SHORT_NAME = SenderReceiverInterface.find(f".//{AB_DIR}{name_node}").text
                try:
                    if name_node == no_fix_name_node[1]:
                        NO_PREFIX_SHORT_NAME = SenderReceiverInterface.find(f".//{AB_DIR}{no_fix_name_node[0]}/{AB_DIR}{no_fix_name_node[1]}").text
                    else:
                        NO_PREFIX_SHORT_NAME = SenderReceiverInterface.find(f".//{AB_DIR}{no_fix_name_node[1]}").text
                    NO_PREFIX_SHORT_NAME = NO_PREFIX_SHORT_NAME.split('/')[-1]
                    SHORT_NAME = SHORT_NAME.split('/')[-1]
                        
                    if SHORT_NAME[starx_idx:] not in [i + NO_PREFIX_SHORT_NAME for i in pre]:
                        SHORT_NAMES.append(SHORT_NAME)
                        NO_PREFIX_SHORT_NAMES.append(NO_PREFIX_SHORT_NAME)
                        UUID.append(SenderReceiverInterface.get('UUID'))
                except:
                   SHORT_NAMES.append(SHORT_NAME)
                   NO_PREFIX_SHORT_NAMES.append(f'Not defined {no_fix_name_node[1]}')
                   UUID.append(SenderReceiverInterface.get('UUID'))
            except:
                SHORT_NAMES.append(SenderReceiverInterface.find(f".//{AB_DIR}SHORT-NAME").text)
                NO_PREFIX_SHORT_NAMES.append(f'Not defined {name_node}')
                UUID.append(SenderReceiverInterface.get('UUID'))
            
    try:
        sheet_ = wb.sheets.add(sheet_name, after=wb.sheets[-1]) 
        sheet_.range('A1').value = ['INDEX', f'{node_tag}/{name_node}', f'{node_tag}/{node_name_tag}/{no_fix_name_node[1]}', 'UUID']  
        sheet_.range('A1').expand(mode="right").font.bold = True
        sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), SHORT_NAMES, NO_PREFIX_SHORT_NAMES, UUID]
        sheet_.range('A1').expand(mode='table').columns.autofit()
    except Exception as err:
        print(err)
        

def SR_CS_check_name(pre:list, node:list, name_node:str, no_fix_name_node:list, child:str, attrib:str, sheet_name:str, starx_idx=0):
    node_tag = get_initials(node[0].replace('-', ' '))+'/'+get_initials(node[1].replace('-', ' '))
    node_name_tag = get_initials(no_fix_name_node[0].replace('-', ' '))
    
    SHORT_NAMES, NO_PREFIX_SHORT_NAMES, UUID = [], [], []   
    SenderReceiverInterfaces = root.findall(f".//{AB_DIR}{node[0]}/{AB_DIR}{node[1]}")
    for SenderReceiverInterface in SenderReceiverInterfaces:
        ATTRIB = SenderReceiverInterface.find(f'.//{AB_DIR}{child}')
        try:
            if ATTRIB.get('DEST') == attrib:
                try:
                    SHORT_NAME = SenderReceiverInterface.find(f".//{AB_DIR}{name_node}").text
                    try:
                        if name_node == no_fix_name_node[1]:
                            NO_PREFIX_SHORT_NAME = SenderReceiverInterface.find(f".//{AB_DIR}{no_fix_name_node[0]}/{AB_DIR}{no_fix_name_node[1]}").text
                        else:
                            NO_PREFIX_SHORT_NAME = SenderReceiverInterface.find(f".//{AB_DIR}{no_fix_name_node[1]}").text
                        NO_PREFIX_SHORT_NAME = NO_PREFIX_SHORT_NAME.split('/')[-1]
                        SHORT_NAME = SHORT_NAME.split('/')[-1]
                            
                        if SHORT_NAME[starx_idx:] not in [i + NO_PREFIX_SHORT_NAME for i in pre]:
                            SHORT_NAMES.append(SHORT_NAME)
                            NO_PREFIX_SHORT_NAMES.append(NO_PREFIX_SHORT_NAME)
                            UUID.append(SenderReceiverInterface.get('UUID'))
                    except:
                       SHORT_NAMES.append(SHORT_NAME)
                       NO_PREFIX_SHORT_NAMES.append(f'Not defined {no_fix_name_node[1]}')
                       UUID.append(SenderReceiverInterface.get('UUID'))
                except:
                    SHORT_NAMES.append(SenderReceiverInterface.find(f".//{AB_DIR}SHORT-NAME").text)
                    NO_PREFIX_SHORT_NAMES.append(f'Not defined {name_node}')
                    UUID.append(SenderReceiverInterface.get('UUID'))
        except:
            ...
            
    try:
        sheet_ = wb.sheets.add(sheet_name, after=wb.sheets[-1]) 
        sheet_.range('A1').value = ['INDEX', f'{node_tag}/{name_node}', f'{node_tag}/{node_name_tag}/{no_fix_name_node[1]}', 'UUID']  
        sheet_.range('A1').expand(mode="right").font.bold = True
        sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), SHORT_NAMES, NO_PREFIX_SHORT_NAMES, UUID]
        sheet_.range('A1').expand(mode='table').columns.autofit()
    except Exception as err:
        print(err)
        
        
def check_name_prefix(pre:str, node:list, name_node:str, sheet_name:str):
    '''
    检查节点text是否以pre开始

    Parameters
    ----------
    pre : str
        前缀的字符.
    node : list
        父节点路径，列表形式[父节点的父节点，父节点].
    name_node : str
        要检查的节点tag.
    sheet_name : str
        新建的子表名.

    Returns
    -------
    None.

    '''
    node_tag = get_initials(node[0].replace('-', ' '))+'/'+get_initials(node[1].replace('-', ' '))
    SHORT_NAMES, UUID = [], []  
    SenderReceiverInterfaces = root.findall(f".//{AB_DIR}{node[0]}/{AB_DIR}{node[1]}")
    for SenderReceiverInterface in SenderReceiverInterfaces:
        SHORT_NAME = SenderReceiverInterface.find(f"{AB_DIR}{name_node}").text
        if not SHORT_NAME.startswith('IB_'):
            SHORT_NAMES.append(SHORT_NAME)
            UUID.append(SenderReceiverInterface.get('UUID'))
    try:
        sheet_ = wb.sheets.add(sheet_name, after=wb.sheets[-1]) 
        sheet_.range('A1').value = ['INDEX', f'{node_tag}/{name_node}', 'UUID']  
        sheet_.range('A1').expand(mode="right").font.bold = True
        sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), SHORT_NAMES, UUID]
        sheet_.range('A1').expand(mode='table').columns.autofit()
    except Exception as err:
        print(err)
 
def check_only_one(node:list, tag_node:str, sheet_name:str):
    '''
    检查父节点下是否仅定义一个tag_node

    Parameters
    ----------
    node : list
        父节点路径，列表形式[父节点的父节点，父节点].
    tag_node : str
        要检查的节点tag.
    sheet_name : str
        新建的子表名.

    Returns
    -------
    None.

    '''
    node_tag = get_initials(node[0].replace('-', ' '))+'/'+get_initials(node[1].replace('-', ' '))
    NUM_NAME_NODE, UUID, SHORT_NAME = [], [], []
    SenderReceiverInterfaces = root.findall(f".//{AB_DIR}{node[0]}/{AB_DIR}{node[1]}")
    for SenderReceiverInterface in SenderReceiverInterfaces:
        num = len(SenderReceiverInterface.findall(f'{AB_DIR}{tag_node}'))
        if num != 1:
            NUM_NAME_NODE.append(num)
            UUID.append(SenderReceiverInterface.get('UUID'))
            SHORT_NAME.append(SenderReceiverInterface.find(f'.//{AB_DIR}SHORT-NAME').text)
    try:
        sheet_ = wb.sheets.add(sheet_name, after=wb.sheets[-1]) 
        sheet_.range('A1').value = ['INDEX', f'{node_tag}/{tag_node}(NUM)', f'{node_tag}/SHORT-NAME', 'UUID']  
        sheet_.range('A1').expand(mode="right").font.bold = True
        sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), NUM_NAME_NODE, SHORT_NAME, UUID]
        sheet_.range('A1').expand(mode='table').columns.autofit()
    except Exception as err:
        print(err)
        
        
def check_only_node(node:list, tag_node:str, sheet_name:str):
    '''
    检查父节点下是否仅定义tag_node

    Parameters
    ----------
    node : list
        父节点路径，列表形式[父节点的父节点，父节点].
    tag_node : str
        要检查的节点tag.
    sheet_name : str
        新建的子表名.

    Returns
    -------
    None.

    '''
    if len(node) == 2:
        node_tag = get_initials(node[0].replace('-', ' '))+'/'+get_initials(node[1].replace('-', ' '))
    elif len(node) == 1:
        node_tag = get_initials(node[0].replace('-', ' '))
    else:
        return None        
    NUM_NAME_NODE, UUID, SHORT_NAME = [], [], []
    RootTop = root.findall(f".//{AB_DIR}{node[0]}")
    for SenderReceiverInterface in RootTop:
        RootBellow = SenderReceiverInterface.find(f"{AB_DIR}{node[1]}")
        ShortName = SenderReceiverInterface.find(f"{AB_DIR}SHORT-NAME").text
        try:
            for child in RootBellow.getchildren():
                chlid_name = child.tag.split('}')[-1]
                if chlid_name != tag_node:
                    NUM_NAME_NODE.append(f"Find {chlid_name}")
                    SHORT_NAME.append(ShortName)
                    UUID.append(child.get('UUID'))
                    break
        except:
            NUM_NAME_NODE.append("Undefined SERVER-CALL-POINTS")
            SHORT_NAME.append(ShortName)
            UUID.append(SenderReceiverInterface.get('UUID'))
    try:
        sheet_ = wb.sheets.add(sheet_name, after=wb.sheets[-1]) 
        sheet_.range('A1').value = ['INDEX', f'{node_tag}/{tag_node}', f'{node_tag}/SHORT-NAME', 'UUID']  
        sheet_.range('A1').expand(mode="right").font.bold = True
        sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), NUM_NAME_NODE, SHORT_NAME, UUID]
        sheet_.range('A1').expand(mode='table').columns.autofit()
    except Exception as err:
        print(err)
        
def check_only_node_v2(node:list, tag_node:str, name:str, sheet_name:str):
    node_tag = get_initials(node[0].replace('-', ' '))
    NUM_NAME_NODE, UUID, SHORT_NAME = [], [], []
    RootTop = root.findall(f".//{AB_DIR}{node[0]}")
    for SenderReceiverInterface in RootTop:
        ShortName = SenderReceiverInterface.find(f".//{AB_DIR}{name}").text
        try:
            for child in SenderReceiverInterface.getchildren():
                chlid_name = child.tag.split('}')[-1]
                if chlid_name != tag_node:
                    NUM_NAME_NODE.append(f"Find {chlid_name}")
                    SHORT_NAME.append(ShortName)
                    UUID.append(child.get('UUID'))
                    break
        except:
            NUM_NAME_NODE.append("Undefined SERVER-CALL-POINTS")
            SHORT_NAME.append(ShortName)
            UUID.append(SenderReceiverInterface.get('UUID'))
    try:
        sheet_ = wb.sheets.add(sheet_name, after=wb.sheets[-1]) 
        sheet_.range('A1').value = ['INDEX', f'{node_tag}/{tag_node}', f'{node_tag}/{name}', 'UUID']  
        sheet_.range('A1').expand(mode="right").font.bold = True
        sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), NUM_NAME_NODE, SHORT_NAME, UUID]
        sheet_.range('A1').expand(mode='table').columns.autofit()
    except Exception as err:
        print(err)
 

def check_config_value(node:list, config_node:str, value:str, name:str, sheet_name:str):
    '''
    检查节点属性配置的值是否正确

    Parameters
    ----------
    node : list
        父节点路径，列表形式[父节点的父节点，父节点].
    config_node : str
        要检查的节点属性.
    value : str
        节点在config_node属性配置的值.
    sheet_name : str
        新建的子表名.

    Returns
    -------
    None.

    '''
    if len(node) == 2:
        SenderReceiverInterfaces = root.findall(f".//{AB_DIR}{node[0]}/{AB_DIR}{node[1]}")
        node_tag = get_initials(node[0].replace('-', ' '))+'/'+get_initials(node[1].replace('-', ' '))
    elif len(node) == 1:
        SenderReceiverInterfaces = root.findall(f".//{AB_DIR}{node[0]}")
        node_tag = get_initials(node[0].replace('-', ' '))
    else:
        return None
    CONFIG_VALUE, UUID, SHORT_NAME = [], [], []
    for SenderReceiverInterface in SenderReceiverInterfaces:
        CONFIG_NODE = SenderReceiverInterface.find(f'.//{AB_DIR}{config_node}')
        try:
            if CONFIG_NODE.text  != value:
                CONFIG_VALUE.append(CONFIG_NODE.text)
                UUID.append(SenderReceiverInterface.get('UUID'))
                SHORT_NAME.append(SenderReceiverInterface.find(f'.//{AB_DIR}{name}').text)
            else:
                ...
        except:
            CONFIG_VALUE.append(f'Undefined {config_node}')
            UUID.append(SenderReceiverInterface.get('UUID'))
            SHORT_NAME.append(SenderReceiverInterface.find(f'.//{AB_DIR}{name}').text)
    try:
        sheet_ = wb.sheets.add(sheet_name, after=wb.sheets[-1]) 
        sheet_.range('A1').value = ['INDEX', f'{node_tag}/{config_node}', f'{node_tag}/{name}', 'UUID']  
        sheet_.range('A1').expand(mode="right").font.bold = True
        sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), CONFIG_VALUE, SHORT_NAME, UUID]
        sheet_.range('A1').expand(mode='table').columns.autofit()
    except Exception as err:
        print(err)
   
    
def check_undefine_node(node:list, tag_node:str, name:str, sheet_name:str):
    '''
    检查节点下是否配置了属性，若配置则报出

    Parameters
    ----------
    node : list
        父节点路径，列表形式[父节点的父节点，父节点].
    tag_node : str
        要检查的节点属性.
    sheet_name : str
        新建的子表名.

    Returns
    -------
    None.

    '''
    if len(node) == 2:
        SenderReceiverInterfaces = root.findall(f".//{AB_DIR}{node[0]}/{AB_DIR}{node[1]}")
        node_tag = get_initials(node[0].replace('-', ' '))+'/'+get_initials(node[1].replace('-', ' '))
    elif len(node) == 1:
        SenderReceiverInterfaces = root.findall(f".//{AB_DIR}{node[0]}")
        node_tag = get_initials(node[0].replace('-', ' '))
    else:
        return None
    
    INVALID_POLICY, SHORT_NAME, UUID = [], [], []
    for SenderReceiverInterface in SenderReceiverInterfaces:
        InvalidationPolicy = SenderReceiverInterface.find(f".//{AB_DIR}{tag_node}")
        if InvalidationPolicy:
            INVALID_POLICY.append(f'Defined {tag_node}')
            SHORT_NAME.append(SenderReceiverInterface.find(f'.//{AB_DIR}{name}').text)
            UUID.append(SenderReceiverInterface.get('UUID'))
    try:
        sheet_ = wb.sheets.add(sheet_name, after=wb.sheets[-1]) 
        sheet_.range('A1').value = ['INDEX', f'{node_tag}/{tag_node}', f'{node_tag}/{name}', 'UUID']  
        sheet_.range('A1').expand(mode="right").font.bold = True
        sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), INVALID_POLICY, SHORT_NAME, UUID]
        sheet_.range('A1').expand(mode='table').columns.autofit()
    except Exception as err:
        print(err)
 
def check_source_and_if_config(node:list, node_source:list, name_node_source:list,
                               pre:str, name_node:str, check_tag:str,
                               check_dir_tag:list, start_idx:int, sheet_name:str):
    '''
    检查node下的属性text是否是以node_source中属性添加pre作为前缀；
    同时检查node下配置的check_tag是否设置check_dir_tag

    Parameters
    ----------
    node : list
        要检查的父节点路径，列表形式[父节点的父节点，父节点].
    node_source : list
        提供来源的父节点路径，列表形式[父节点的父节点，父节点].
    name_node_source : list
        提供来源的节点路径.
    pre : str
        检查的节点text是以pre做前缀的提供来源的节点text.
    name_node : str
        要检查的节点.
    check_tag : str
        要检查的父节点下的某一属性祖父节点.
    check_dir_tag : list
        检查的父节点下的某一属性下的路径，形式[父节点，属性].
    sheet_name : str
        新建的子表名.

    Returns
    -------
    None.

    '''
    node_tag = get_initials(node[0].replace('-', ' '))+'/'+get_initials(node[1].replace('-', ' '))
    node_source_tag = get_initials(node_source[0].replace('-', ' '))+'/'+get_initials(node_source[1].replace('-', ' '))
    NODE = root.findall(f".//{AB_DIR}{node[0]}/{AB_DIR}{node[1]}")
    SOURCE_NODE = root.findall(f".//{AB_DIR}{node_source[0]}/{AB_DIR}{node_source[1]}")
    SOURCES, UUID, ERROR, CS_SHORT_NAMES = [], [], [], []
    for source_ in SOURCE_NODE:
        SOURCE = source_.find(f'.//{AB_DIR}{name_node_source}').text
        SOURCES.append(pre + SOURCE[start_idx:])
    for node_ in NODE:
        NODE_NAME = node_.find(f'.//{AB_DIR}{name_node[0]}/{AB_DIR}{name_node[1]}').text
        if NODE_NAME not in SOURCES:
            UUID.append(node_.get('UUID'))
            CS_SHORT_NAMES.append(NODE_NAME)
            ERROR.append(f'{NODE_NAME} not in {node_source_tag}')
        else:
            ind = SOURCES.index(NODE_NAME)
            type_ref = (node_.find(f'.//{AB_DIR}TYPE-TREF')).get('DEST')
            if type_ref == 'APPLICATION-PRIMITIVE-DATA-TYPE':
                param = SOURCE_NODE[ind].find(f'.//{AB_DIR}VALUE-SPEC//{AB_DIR}APPLICATION-VALUE-SPECIFICATION')
                Short_label = SOURCE_NODE[ind].find(f'.//{AB_DIR}SHORT-LABEL').text
                Flag = param.find(f'.//{AB_DIR}SW-VALUES-PHYS')
                if not Flag:
                    ERROR.append(f'{Short_label} undefined PHYS(P)')
                    UUID.append(SOURCE_NODE[ind].get('UUID'))
                    CS_SHORT_NAMES.append(NODE_NAME)
            elif type_ref == 'APPLICATION-RECORD-DATA-TYPE':
                param = SOURCE_NODE[ind].findall(f'.//{AB_DIR}VALUE-SPEC/{AB_DIR}RECORD-VALUE-SPECIFICATION//{AB_DIR}APPLICATION-VALUE-SPECIFICATION')
                Error_info  = 'R_'
                IF_error = 0
                for i,j in enumerate(param):
                    Flag = j.find(f'.//{AB_DIR}SW-VALUES-PHYS')
                    if not Flag:
                        Short_label = j.find(f'.//{AB_DIR}SHORT-LABEL').text
                        Error_info += f'{Short_label} undefined PHYS; '
                        IF_error = 1
                if IF_error:
                    ERROR.append(Error_info)
                    UUID.append(SOURCE_NODE[ind].get('UUID'))
                    CS_SHORT_NAMES.append(NODE_NAME)
            elif type_ref == 'APPLICATION-ARRAY-DATA-TYPE':
                param = SOURCE_NODE[ind].findall(f'.//{AB_DIR}VALUE-SPEC/{AB_DIR}ARRAY-VALUE-SPECIFICATION//{AB_DIR}APPLICATION-VALUE-SPECIFICATION')
                Error_info  = 'A_'
                for i,j in enumerate(param):
                    Flag = j.find(f'.//{AB_DIR}SW-VALUES-PHYS')
                    if not Flag:
                        Short_label = j.find(f'.//{AB_DIR}SHORT-LABEL').text
                        Error_info += f'{Short_label} undefined PHYS'
                        IF_error = 1
                if IF_error:
                    UUID.append(SOURCE_NODE[ind].get('UUID'))
                    CS_SHORT_NAMES.append(NODE_NAME)
                    ERROR.append(Error_info)
    try:           
        sheet_ = wb.sheets.add(sheet_name, after=wb.sheets[-1]) 
        sheet_.range('A1').value = ['INDEX', 'ERROR', f'{node_tag}/{name_node}', 'UUID']  
        sheet_.range('A1').expand(mode="right").font.bold = True
        sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), ERROR, CS_SHORT_NAMES, UUID]
        sheet_.range('A1').expand(mode='table').columns.autofit()
    except Exception as err:
        print(err)
        

def check_source_and_config(node:list, node_source:list, name_node_source:list,
                               pre:str, name_node:str, check_node:str, config_value:str, sheet_name:str):
    '''
    检查node下的属性text是否是以node_source中属性添加pre作为前缀；
    同时检查node下配置的check_tag是否设置check_dir_tag

    Parameters
    ----------
    node : list
        要检查的父节点路径，列表形式[父节点的父节点，父节点].
    node_source : list
        提供来源的父节点路径，列表形式[父节点的父节点，父节点].
    name_node_source : list
        提供来源的节点路径.
    pre : str
        检查的节点text是以pre做前缀的提供来源的节点text.
    name_node : str
        要检查的节点.
    check_tag : str
        要检查的父节点下的某一属性祖父节点.
    check_dir_tag : list
        检查的父节点下的某一属性下的路径，形式[父节点，属性].
    sheet_name : str
        新建的子表名.

    Returns
    -------
    None.

    '''
    node_tag = get_initials(node[0].replace('-', ' '))+'/'+get_initials(node[1].replace('-', ' '))
    node_source_tag = get_initials(node_source[0].replace('-', ' '))+'/'+get_initials(node_source[1].replace('-', ' '))
    NODE = root.findall(f".//{AB_DIR}{node[0]}/{AB_DIR}{node[1]}")
    SOURCE_NODE = root.findall(f".//{AB_DIR}{node_source[0]}/{AB_DIR}{node_source[1]}")
    SOURCE_INFOS, UUID, ERROR, CS_SHORT_NAMES = [], [], [], []
    for source_ in SOURCE_NODE:
        SOURCE_INFO = source_.find(f'.//{AB_DIR}{name_node_source}').text
        SOURCE_INFOS.append(SOURCE_INFO)
                
    for node_ in NODE:
        DESTINATION_INFO = node_.find(f'{AB_DIR}{name_node}').text
        if pre + DESTINATION_INFO not in SOURCE_INFOS:
            UUID.append(node_.get('UUID'))
            CS_SHORT_NAMES.append(DESTINATION_INFO)
            ERROR.append(f'{DESTINATION_INFO} not in {node_source_tag}')
        else:
            ind = SOURCE_INFOS.index(pre + DESTINATION_INFO)
            FLAG = (SOURCE_NODE[ind].find(f'.//{AB_DIR}{check_node}').text).split('/')[-1]
            if FLAG not in config_value:
                UUID.append(source_.get('UUID'))
                ERROR.append(f"DataType: {FLAG}")
                CS_SHORT_NAMES.append(SOURCE_INFO)
    try:           
        sheet_ = wb.sheets.add(sheet_name, after=wb.sheets[-1]) 
        sheet_.range('A1').value = ['INDEX', 'ERROR', f'{node_tag}/{name_node}', 'UUID']  
        sheet_.range('A1').expand(mode="right").font.bold = True
        sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), ERROR, CS_SHORT_NAMES, UUID]
        sheet_.range('A1').expand(mode='table').columns.autofit()
    except Exception as err:
        print(err)
        

def check_source_with_prefix_and_config(node:list, node_source:list, name_node_source:list,
                               pre:str, start_index:int, name_node:str, check_node:str,
                               config_value:str, sheet_name:str):
    node_tag = get_initials(node[0].replace('-', ' '))+'/'+get_initials(node[1].replace('-', ' '))
    node_source_tag = get_initials(node_source[0].replace('-', ' '))+'/'+get_initials(node_source[1].replace('-', ' '))
    NODE = root.findall(f".//{AB_DIR}{node[0]}//{AB_DIR}{node[1]}")
    SOURCE_NODE = root.findall(f".//{AB_DIR}{node_source[0]}/{AB_DIR}{node_source[1]}")
    SOURCE_INFOS, UUID, ERROR, CS_SHORT_NAMES = [], [], [], []
    for source_ in SOURCE_NODE:
        SOURCE_INFO = source_.find(f'.//{AB_DIR}{name_node_source}').text
        SOURCE_INFOS.append(pre + SOURCE_INFO[start_index:])
                
    for node_ in NODE:
        DESTINATION_INFO = node_.find(f'.//{AB_DIR}{name_node}').text
        if DESTINATION_INFO not in SOURCE_INFOS:
            UUID.append(node_.get('UUID'))
            CS_SHORT_NAMES.append(DESTINATION_INFO)
            ERROR.append(f'{DESTINATION_INFO} not in {node_source_tag}')
        else:
            ind = SOURCE_INFOS.index(DESTINATION_INFO)
            FLAG = (SOURCE_NODE[ind].find(f'.//{AB_DIR}{check_node}').text).split('/')[-1]
            if FLAG not in config_value:
                UUID.append(source_.get('UUID'))
                ERROR.append(f"{check_node}: {FLAG}")
                CS_SHORT_NAMES.append(DESTINATION_INFO)
    try:           
        sheet_ = wb.sheets.add(sheet_name, after=wb.sheets[-1]) 
        sheet_.range('A1').value = ['INDEX', 'ERROR', f'{node_tag}/{name_node}', 'UUID']  
        sheet_.range('A1').expand(mode="right").font.bold = True
        sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), ERROR, CS_SHORT_NAMES, UUID]
        sheet_.range('A1').expand(mode='table').columns.autofit()
    except Exception as err:
        print(err)
        
def check_source_and_dest(node:list, node_source:list, name_node_source:list, name_node:str, check_dest:str, sheet_name:str):
    '''
    

    Parameters
    ----------
    node : list
        要检查的父节点路径，列表形式[父节点的父节点，父节点].
    node_source : list
        提供来源的父节点路径，列表形式[父节点的父节点，父节点].
    name_node_source : list
        提供来源的节点路径.
    name_node : str
        要检查的节点..
    check_dest : str
        检查的DEST属性值.
    sheet_name : str
        新建的子表名.

    Returns
    -------
    None.

    '''
    node_tag = get_initials(node[0].replace('-', ' '))+'/'+get_initials(node[1].replace('-', ' '))
    node_source_tag = get_initials(node_source[0].replace('-', ' '))+'/'+get_initials(node_source[1].replace('-', ' '))
    NODE = root.findall(f".//{AB_DIR}{node[0]}/{AB_DIR}{node[1]}")
    SOURCE_NODE = root.findall(f".//{AB_DIR}{node_source[0]}//{AB_DIR}{node_source[1]}")
    SR_DATAELEMENT_SHORT_NAMES, UUID, ERROR, CS_SHORT_NAMES = [], [], [], []
    for source_ in SOURCE_NODE:
        if source_.get('DEST') == check_dest:
            IBDatatypeRef = source_.text
            IBDatatypeRef = IBDatatypeRef.split('/')[-1]
            SR_DATAELEMENT_SHORT_NAMES.append(IBDatatypeRef)
    for node_ in NODE:
        CONSTANT_SPECIFICATION_SHORT_NAME = node_.find(f'{AB_DIR}{name_node}').text
        if CONSTANT_SPECIFICATION_SHORT_NAME not in SR_DATAELEMENT_SHORT_NAMES:
            UUID.append(node_.get('UUID'))
            CS_SHORT_NAMES.append(CONSTANT_SPECIFICATION_SHORT_NAME)
            ERROR.append(f'{CONSTANT_SPECIFICATION_SHORT_NAME} not in {node_source_tag}/{name_node_source[1]}')
    try:           
        sheet_ = wb.sheets.add(sheet_name, after=wb.sheets[-1]) 
        sheet_.range('A1').value = ['INDEX', 'ERROR', f'{node_tag}/{name_node}', 'UUID']  
        sheet_.range('A1').expand(mode="right").font.bold = True
        sheet_.range('A2').options(transpose=True).value = [list(range(1,len(UUID)+1)), ERROR, CS_SHORT_NAMES, UUID]
        sheet_.range('A1').expand(mode='table').columns.autofit()
    except Exception as err:
        print(err)
        
#%%主函数
if __name__ == '__main__':
    global AB_DIR
    AB_DIR = '{http://autosar.org/schema/r4.0}'
    FILE_NAME = '0702htpcArxml'
    
    app = xw.App(visible=False, add_book=False)
    
    file_path = 'D:/Hua/My_Own_Utilities/arxml/arxml_py/'+FILE_NAME+'.arxml'
    
    tree = ET.parse(file_path)
    root = tree.getroot()
    #%%SR_Interface
    ##%%N003(2)
    EXCEL_PATH = './ARXML/'+ FILE_NAME+'_SR_Interface.xlsx' 
    try:
        wb = app.books.open(EXCEL_PATH)
    except:
        wb = app.books.add()
    pre = ['IF_']
    node = ['ELEMENTS', 'SENDER-RECEIVER-INTERFACE']    
    name_node = 'SHORT-NAME'
    no_fix_name_node = ['VARIABLE-DATA-PROTOTYPE', 'SHORT-NAME']
    sheet_name = 'N003(2)'
    check_name(pre, node, name_node, no_fix_name_node, sheet_name)
    ##%%A006(1)
    node = ['ELEMENTS', 'SENDER-RECEIVER-INTERFACE']    
    tag_node = 'DATA-ELEMENTS'
    sheet_name = 'A006(1)'
    check_only_one(node, tag_node, sheet_name)
    ##%% N019(1)
    pre = ['APDT_', 'AADT_', 'ARDT_']
    node = ['DATA-ELEMENTS', 'VARIABLE-DATA-PROTOTYPE']    
    name_node = 'TYPE-TREF'
    no_fix_name_node = ['VARIABLE-DATA-PROTOTYPE', 'SHORT-NAME']
    sheet_name = 'N019(1)'
    check_name(pre, node, name_node, no_fix_name_node, sheet_name)
    ##%%A041(1)
    node = ['SENDER-RECEIVER-INTERFACE', 'DATA-ELEMENTS']
    config_node = 'SW-CALIBRATION-ACCESS'
    value = 'READ-ONLY'
    sheet_name = 'A041(1)'
    check_config_value(node, config_node, value, 'SHORT-NAME', sheet_name)
    ##%%A042(1)
    node = ['ELEMENTS', 'SENDER-RECEIVER-INTERFACE']
    config_node = 'SW-IMPL-POLICY'
    value = 'STANDARD'
    sheet_name = 'A042(1)'
    check_config_value(node, config_node, value, 'SHORT-NAME', sheet_name)
    ##%%A111(1)
    node = ['SENDER-RECEIVER-INTERFACE', 'DATA-ELEMENTS']
    node_source = ['ELEMENTS', 'CONSTANT-SPECIFICATION']
    name_node_source = 'SHORT-NAME'
    name_node = ['VARIABLE-DATA-PROTOTYPE', 'SHORT-NAME']
    pre = ''
    start_idx = 3
    check_tag = 'APPLICATION-VALUE-SPECIFICATION'
    check_dir_tag = ['SW-VALUE-CONT', 'SW-VALUES-PHYS']
    sheet_name = 'A111(1)'
    check_source_and_if_config(node, node_source, name_node_source, pre, name_node, check_tag, check_dir_tag, start_idx, sheet_name)
    ##%%A131(1)
    pre = ['IV_']
    node = ['SENDER-RECEIVER-INTERFACE', 'DATA-ELEMENTS']
    name_node = 'CONSTANT-REF'
    no_fix_name_node = ['VARIABLE-DATA-PROTOTYPE', 'SHORT-NAME']
    sheet_name = 'A131(1)'
    check_name(pre, node, name_node, no_fix_name_node, sheet_name)
    ##%%A067(1)
    node = ['ELEMENTS', 'SENDER-RECEIVER-INTERFACE']
    tag_node = 'INVALIDATION-POLICY'
    name = 'SHORT-NAME'
    sheet_name = 'A067(1)'
    check_undefine_node(node, tag_node, name, sheet_name)
    
    wb.sheets[0].delete()
    wb.save()
    weight_file = Path.cwd() / wb.name
    new_name =  Path.cwd() / EXCEL_PATH
    wb.close()
    try:
        weight_file.rename(new_name)
    except:
        ...
    
    #%%SR_Ports
    EXCEL_PATH = './ARXML/'+ FILE_NAME+'_SR_Ports.xlsx' 
    try:
        wb = app.books.open(EXCEL_PATH)
    except:
        wb = app.books.add()
    ##%%N007(1)
    pre = ['R_']
    node = ['PORTS', 'R-PORT-PROTOTYPE']
    name_node = 'SHORT-NAME'
    no_fix_name_node = ['NONQUEUED-RECEIVER-COM-SPEC', 'DATA-ELEMENT-REF']
    child = 'REQUIRED-INTERFACE-TREF'
    attrib = 'SENDER-RECEIVER-INTERFACE'
    sheet_name = 'N007(1)'
    SR_port_check_name(pre, node, name_node, no_fix_name_node, child, attrib, sheet_name)
    ##%%N008(1)
    pre = ['P_']
    node = ['PORTS', 'P-PORT-PROTOTYPE']
    name_node = 'SHORT-NAME'
    no_fix_name_node = ['NONQUEUED-SENDER-COM-SPEC', 'DATA-ELEMENT-REF']
    child = 'PROVIDED-INTERFACE-TREF'
    attrib = 'SENDER-RECEIVER-INTERFACE'
    sheet_name = 'N008(1)'
    SR_port_check_name(pre, node, name_node, no_fix_name_node, child, attrib, sheet_name)
    
    wb.sheets[0].delete()
    wb.save()
    weight_file = Path.cwd() / wb.name
    new_name =  Path.cwd() / EXCEL_PATH
    wb.close()
    try:
        weight_file.rename(new_name)
    except:
        ...
    
    #%%SRPorts_ComSpec
    EXCEL_PATH = './ARXML/'+ FILE_NAME+'_SRPorts_ComSpec.xlsx' 
    try:
        wb = app.books.open(EXCEL_PATH)
    except:
        wb = app.books.add()
    ##%%A030(2)
    pre = ['IV_']
    node = ['PORTS', 'P-PORT-PROTOTYPE']
    name_node = 'CONSTANT-REF'
    no_fix_name_node = ['NONQUEUED-SENDER-COM-SPEC', 'DATA-ELEMENT-REF']
    child = 'PROVIDED-INTERFACE-TREF'
    attrib = 'SENDER-RECEIVER-INTERFACE'
    sheet_name = 'A030(2)'
    SR_CS_check_name(pre, node, name_node, no_fix_name_node, child, attrib, sheet_name)
    ##%%A031(2)
    pre = ['IV_']
    node = ['PORTS', 'R-PORT-PROTOTYPE']
    name_node = 'CONSTANT-REF'
    child = 'REQUIRED-INTERFACE-TREF'
    attrib = 'SENDER-RECEIVER-INTERFACE'
    no_fix_name_node = ['NONQUEUED-RECEIVER-COM-SPEC', 'DATA-ELEMENT-REF']
    sheet_name = 'A031(2)'
    SR_CS_check_name(pre, node, name_node, no_fix_name_node, child, attrib, sheet_name)
    ##%%A032(3)
    node = node = ['NONQUEUED-RECEIVER-COM-SPEC']#
    config_node = 'ALIVE-TIMEOUT'
    value = '0'
    sheet_name = 'A032(3)'
    check_config_value(node, config_node, value, 'DATA-ELEMENT-REF', sheet_name)
    ##%%A116(0)
    node = node = ['NONQUEUED-RECEIVER-COM-SPEC']
    config_node = 'HANDLE-OUT-OF-RANGE'
    value = 'NONE'
    sheet_name = 'A116(0)'
    check_config_value(node, config_node, value, 'DATA-ELEMENT-REF', sheet_name)
    ##%%A117(0)
    node = node = ['NONQUEUED-RECEIVER-COM-SPEC']
    config_node = 'ENABLE-UPDATE'
    value = 'false'
    sheet_name = 'A117(0)'
    check_config_value(node, config_node, value, 'DATA-ELEMENT-REF', sheet_name)
    ##%%A118(0)
    node = node = ['NONQUEUED-RECEIVER-COM-SPEC']
    config_node = 'HANDLE-NEVER-RECEIVED'
    value = 'false'
    sheet_name = 'A118(0)'
    check_config_value(node, config_node, value, 'DATA-ELEMENT-REF', sheet_name)
    ##%%A119(0)
    node = node = ['NONQUEUED-RECEIVER-COM-SPEC']
    config_node = 'HANDLE-TIMEOUT-TYPE'
    value = 'NONE'
    sheet_name = 'A119(0)'
    check_config_value(node, config_node, value, 'DATA-ELEMENT-REF', sheet_name)
    ##%%A120(0)
    node = node = ['NONQUEUED-RECEIVER-COM-SPEC']
    config_node = 'USES-END-TO-END-PROTECTION'
    value = 'false'
    sheet_name = 'A120(0)'
    check_config_value(node, config_node, value, 'DATA-ELEMENT-REF', sheet_name)
    ##%%A122(0)
    node = node = ['NONQUEUED-SENDER-COM-SPEC']
    config_node = 'USES-END-TO-END-PROTECTION'
    value = 'false'
    sheet_name = 'A122(0)'
    check_config_value(node, config_node, value, 'DATA-ELEMENT-REF', sheet_name)
    ##%%A123(0)
    node = node = ['NONQUEUED-SENDER-COM-SPEC']
    config_node = 'HANDLE-OUT-OF-RANGE'
    value = 'NONE'
    sheet_name = 'A123(0)'
    check_config_value(node, config_node, value, 'DATA-ELEMENT-REF', sheet_name)
    ##%%A033(3)
    node = ['NONQUEUED-SENDER-COM-SPEC']
    tag_node = 'TRANSMISSION-ACKNOWLEDGE'
    sheet_name = 'A033(3)_TRANSMISSION'
    check_undefine_node(node, tag_node, 'DATA-ELEMENT-REF', sheet_name)
    
    tag_node = 'NEYWORK-REPRESENTATION'
    sheet_name = 'A033(3)_NEYWORK'
    check_undefine_node(node, tag_node, 'DATA-ELEMENT-REF', sheet_name)
    
    node = ['COMPOSITION-SW-COMPONENT-TYPE']
    tag_node = 'NEYWORK-REPRESENTATION'
    sheet_name = 'A033(3)_CSCT_NEYWORK'
    check_undefine_node(node, tag_node, 'DATA-ELEMENT-REF', sheet_name)
    
    wb.sheets[0].delete()
    wb.save()
    weight_file = Path.cwd() / wb.name
    new_name =  Path.cwd() / EXCEL_PATH
    wb.close()
    try:
        weight_file.rename(new_name)
    except:
        ...
    
    #%%InterBehavior
    EXCEL_PATH = './ARXML/'+ FILE_NAME+'_InterBehavior.xlsx' 
    try:
        wb = app.books.open(EXCEL_PATH)
    except:
        wb = app.books.add()
    ##%% N010(3)
    pre = 'IB_'
    node = ['INTERNAL-BEHAVIORS', 'SWC-INTERNAL-BEHAVIOR']    
    name_node = 'SHORT-NAME'
    sheet_name = 'N010(3)'
    check_name_prefix(pre, node, name_node, sheet_name)
    ##%% A125(0)
    node = ['ELEMENTS', 'DATA-TYPE-MAPPING-SET']  
    node_source = ['SWC-INTERNAL-BEHAVIOR', 'DATA-TYPE-MAPPING-REF']  
    name_node = 'SHORT-NAME'
    name_node_source = ['DATA-TYPE-MAPPING-REFS', 'DATA-TYPE-MAPPING-REF']
    check_dest = 'DATA-TYPE-MAPPING-SET'
    sheet_name = 'A125(0)'
    check_source_and_dest(node, node_source, name_node_source, name_node, check_dest, sheet_name)
    ##%%A139_1(0)
    node = node = ['INTERNAL-BEHAVIORS', 'SWC-INTERNAL-BEHAVIOR']
    config_node = 'HANDLE-TERMINATION-AND-RESTART'
    value = 'NO-SUPPORT'
    sheet_name = 'A139_1(0)'
    check_config_value(node, config_node, value, 'SHORT-NAME', sheet_name)
    ##%%A139_2(0)
    node = node = ['INTERNAL-BEHAVIORS', 'SWC-INTERNAL-BEHAVIOR']
    config_node = 'SUPPORTS-MULTIPLE-INSTANTIATION'
    value = 'false'
    sheet_name = 'A139_2(0)'
    check_config_value(node, config_node, value, 'SHORT-NAME', sheet_name)
    ##%%A017(3)
    pre = ['IMP_']
    start_index = 3
    node = ['ELEMENTS', 'SWC-IMPLEMENTATION']
    name_node = 'SHORT-NAME'
    no_fix_name_node = ['SWC-IMPLEMENTATION', 'BEHAVIOR-REF']
    sheet_name = 'A017(3)'
    check_name_IB(pre, node, name_node, no_fix_name_node, sheet_name, start_index)
    ##%%A020(1)
    node = ['ELEMENTS', 'SWC-IMPLEMENTATION']
    config_node = 'PROGRAMMING-LANGUAGE'
    value = 'C'
    sheet_name = 'A020(1)'
    check_config_value(node, config_node, value, 'SHORT-NAME', sheet_name)
    
    wb.sheets[0].delete()
    wb.save()
    weight_file = Path.cwd() / wb.name
    new_name =  Path.cwd() / EXCEL_PATH
    wb.close()
    try:
        weight_file.rename(new_name)
    except:
        ...
    
    #%%Data_access
    EXCEL_PATH = './ARXML/'+ FILE_NAME+'_Data_access.xlsx' 
    try:
        wb = app.books.open(EXCEL_PATH)
    except:
        wb = app.books.add()
    ##N012(4)  
    pre = ['ds_']
    node = ['DATA-SEND-POINTS', 'VARIABLE-ACCESS ']
    name_node = 'SHORT-NAME'
    no_fix_name_node = ['AUTOSAR-VARIABLE-IREF', 'TARGET-DATA-PROTOTYPE-REF']
    sheet_name = 'N012(4)'
    check_name(pre, node, name_node, no_fix_name_node, sheet_name)
    ##N047(2)  
    pre = ['dr_']
    node = ['DATA-RECEIVE-POINT-BY-ARGUMENTS', 'VARIABLE-ACCESS']
    name_node = 'SHORT-NAME'
    no_fix_name_node = ['AUTOSAR-VARIABLE-IREF', 'TARGET-DATA-PROTOTYPE-REF']
    sheet_name = 'N047(2)'
    check_name(pre, node, name_node, no_fix_name_node, sheet_name)
    wb.sheets[0].delete()
    wb.save()
    weight_file = Path.cwd() / wb.name
    new_name =  Path.cwd() / EXCEL_PATH
    wb.close()
    try:
        weight_file.rename(new_name)
    except:
        ...
        
    #%%Operation_access
    EXCEL_PATH = './ARXML/'+ FILE_NAME+'_Operation_access.xlsx' 
    try:
        wb = app.books.open(EXCEL_PATH)
    except:
        wb = app.books.add()
    ##%%A014(4)
    node = ['SERVER-CALL-POINTS']    
    tag_node = 'SYNCHRONOUS-SERVER-CALL-POINT'
    sheet_name = 'A014(4)'
    check_only_node_v2(node, tag_node, 'SHORT-NAME', sheet_name)
    ##%%N013(4)
    pre = ['sc_']
    start_index = 0
    node = ['SYNCHRONOUS-SERVER-CALL-POINT']
    name_node = 'SHORT-NAME'
    no_fix_name_node = ['OPERATION-IREF', 'TARGET-REQUIRED-OPERATION-REF']
    sheet_name = 'N013(4)'
    check_name(pre, node, name_node, no_fix_name_node, sheet_name, start_index)
    ##%%N018(2)
    node = ['SYNCHRONOUS-SERVER-CALL-POINT']
    config_node = 'TIMEOUT'
    value = '0'
    sheet_name = 'N018(2)'
    check_config_value(node, config_node, value, 'SHORT-NAME', sheet_name)
    
    wb.sheets[0].delete()
    wb.save()
    weight_file = Path.cwd() / wb.name
    new_name =  Path.cwd() / EXCEL_PATH
    wb.close()
    try:
        weight_file.rename(new_name)
    except:
        ...
    
    #%%ModeSwitch
    EXCEL_PATH = './ARXML/'+ FILE_NAME+'_ModeSwitch.xlsx' 
    try:
        wb = app.books.open(EXCEL_PATH)
    except:
        wb = app.books.add()
    ##%%A110(0)
    node = ['ELEMENTS', 'MODE-DECLARATION-GROUP']
    node_source = ['ELEMENTS', 'DATA-TYPE-MAPPING-SET']
    name_node_source = 'SHORT-NAME'
    pre = 'DTMS_'
    name_node = 'SHORT-NAME'
    check_node = 'IMPLEMENTATION-DATA-TYPE-REF'
    config_value = ['UInt8']
    sheet_name = 'A110(0)'
    check_source_and_config(node, node_source, name_node_source, pre, name_node, check_node, config_value, sheet_name)
    ##%%A136(0)
    node = ['ELEMENTS', 'MODE-SWITCH-INTERFACE']
    config_node = 'SW-CALIBRATION-ACCESS'
    value = 'READ-ONLY'
    sheet_name = 'A136(0)'
    check_config_value(node, config_node, value, 'SHORT-NAME', sheet_name)
    
    wb.sheets[0].delete()
    wb.save()
    weight_file = Path.cwd() / wb.name
    new_name =  Path.cwd() / EXCEL_PATH
    wb.close()
    try:
        weight_file.rename(new_name)
    except:
        ...
    
    #%%CSWC_Port
    EXCEL_PATH = './ARXML/'+ FILE_NAME+'_CSWC_Port.xlsx' 
    try:
        wb = app.books.open(EXCEL_PATH)
    except:
        wb = app.books.add()
    ##%%A104(0)
    node = ['COMPOSITION-SW-COMPONENT-TYPE', 'P-PORT-PROTOTYPE']
    node_source = ['ELEMENTS', 'SENDER-RECEIVER-INTERFACE']
    name_node_source = 'SHORT-NAME'
    pre = 'P_'
    name_node = 'SHORT-NAME'
    check_node = 'IS-SERVICE'
    config_value = ['false']
    sheet_name = 'A104(0)_PPort'
    start_index = 3
    check_source_with_prefix_and_config(node, node_source, name_node_source, pre, start_index, name_node, check_node, config_value, sheet_name)
    # ##%%A104(0)
    node = ['COMPOSITION-SW-COMPONENT-TYPE', 'R-PORT-PROTOTYPE']
    node_source = ['ELEMENTS', 'SENDER-RECEIVER-INTERFACE']
    name_node_source = 'SHORT-NAME'
    pre = 'R_'
    name_node = 'SHORT-NAME'
    check_node = 'IS-SERVICE'
    config_value = ['false']
    sheet_name = 'A104(0)_RPort'
    start_index = 3
    check_source_with_prefix_and_config(node, node_source, name_node_source, pre, start_index, name_node, check_node, config_value, sheet_name)
    
    wb.sheets[0].delete()
    wb.save()
    weight_file = Path.cwd() / wb.name
    new_name =  Path.cwd() / EXCEL_PATH
    wb.close()
    try:
        weight_file.rename(new_name)
    except:
        ...
    #%%Quit
    app.quit()
        
    #%%

