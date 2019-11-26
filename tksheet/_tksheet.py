from collections import defaultdict, deque
from itertools import islice, repeat, accumulate
from math import floor, ceil
from tkinter import ttk
import bisect
import csv as csv_module
import io
import pickle
import re
import tkinter as tk
import zlib
# for mac bindings
from platform import system as get_os


Color_Map_ = {
    'alice blue': '#F0F8FF',
    'AliceBlue': '#F0F8FF',
    'antique white': '#FAEBD7',
    'AntiqueWhite': '#FAEBD7',
    'AntiqueWhite1': '#FFEFDB',
    'AntiqueWhite2': '#EEDFCC',
    'AntiqueWhite3': '#CDC0B0',
    'AntiqueWhite4': '#8B8378',
    'aquamarine': '#7FFFD4',
    'aquamarine1': '#7FFFD4',
    'aquamarine2': '#76EEC6',
    'aquamarine3': '#66CDAA',
    'aquamarine4': '#458B74',
    'azure': '#F0FFFF',
    'azure1': '#F0FFFF',
    'azure2': '#E0EEEE',
    'azure3': '#C1CDCD',
    'azure4': '#838B8B',
    'beige': '#F5F5DC',
    'bisque': '#FFE4C4',
    'bisque1': '#FFE4C4',
    'bisque2': '#EED5B7',
    'bisque3': '#CDB79E',
    'bisque4': '#8B7D6B',
    'black': '#000000',
    'blanched almond': '#FFEBCD',
    'BlanchedAlmond': '#FFEBCD',
    'blue': '#0000FF',
    'blue violet': '#8A2BE2',
    'blue1': '#0000FF',
    'blue2': '#0000EE',
    'blue3': '#0000CD',
    'blue4': '#00008B',
    'BlueViolet': '#8A2BE2',
    'brown': '#A52A2A',
    'brown1': '#FF4040',
    'brown2': '#EE3B3B',
    'brown3': '#CD3333',
    'brown4': '#8B2323',
    'burlywood': '#DEB887',
    'burlywood1': '#FFD39B',
    'burlywood2': '#EEC591',
    'burlywood3': '#CDAA7D',
    'burlywood4': '#8B7355',
    'cadet blue': '#5F9EA0',
    'CadetBlue': '#5F9EA0',
    'CadetBlue1': '#98F5FF',
    'CadetBlue2': '#8EE5EE',
    'CadetBlue3': '#7AC5CD',
    'CadetBlue4': '#53868B',
    'chartreuse': '#7FFF00',
    'chartreuse1': '#7FFF00',
    'chartreuse2': '#76EE00',
    'chartreuse3': '#66CD00',
    'chartreuse4': '#458B00',
    'chocolate': '#D2691E',
    'chocolate1': '#FF7F24',
    'chocolate2': '#EE7621',
    'chocolate3': '#CD661D',
    'chocolate4': '#8B4513',
    'coral': '#FF7F50',
    'coral1': '#FF7256',
    'coral2': '#EE6A50',
    'coral3': '#CD5B45',
    'coral4': '#8B3E2F',
    'cornflower blue': '#6495ED',
    'CornflowerBlue': '#6495ED',
    'cornsilk': '#FFF8DC',
    'cornsilk1': '#FFF8DC',
    'cornsilk2': '#EEE8CD',
    'cornsilk3': '#CDC8B1',
    'cornsilk4': '#8B8878',
    'cyan': '#00FFFF',
    'cyan1': '#00FFFF',
    'cyan2': '#00EEEE',
    'cyan3': '#00CDCD',
    'cyan4': '#008B8B',
    'dark blue': '#00008B',
    'dark cyan': '#008B8B',
    'dark goldenrod': '#B8860B',
    'dark gray': '#A9A9A9',
    'dark green': '#006400',
    'dark grey': '#A9A9A9',
    'dark khaki': '#BDB76B',
    'dark magenta': '#8B008B',
    'dark olive green': '#556B2F',
    'dark orange': '#FF8C00',
    'dark orchid': '#9932CC',
    'dark red': '#8B0000',
    'dark salmon': '#E9967A',
    'dark sea green': '#8FBC8F',
    'dark slate blue': '#483D8B',
    'dark slate gray': '#2F4F4F',
    'dark slate grey': '#2F4F4F',
    'dark turquoise': '#00CED1',
    'dark violet': '#9400D3',
    'DarkBlue': '#00008B',
    'DarkCyan': '#008B8B',
    'DarkGoldenrod': '#B8860B',
    'DarkGoldenrod1': '#FFB90F',
    'DarkGoldenrod2': '#EEAD0E',
    'DarkGoldenrod3': '#CD950C',
    'DarkGoldenrod4': '#8B6508',
    'DarkGray': '#A9A9A9',
    'DarkGreen': '#006400',
    'DarkGrey': '#A9A9A9',
    'DarkKhaki': '#BDB76B',
    'DarkMagenta': '#8B008B',
    'DarkOliveGreen': '#556B2F',
    'DarkOliveGreen1': '#CAFF70',
    'DarkOliveGreen2': '#BCEE68',
    'DarkOliveGreen3': '#A2CD5A',
    'DarkOliveGreen4': '#6E8B3D',
    'DarkOrange': '#FF8C00',
    'DarkOrange1': '#FF7F00',
    'DarkOrange2': '#EE7600',
    'DarkOrange3': '#CD6600',
    'DarkOrange4': '#8B4500',
    'DarkOrchid': '#9932CC',
    'DarkOrchid1': '#BF3EFF',
    'DarkOrchid2': '#B23AEE',
    'DarkOrchid3': '#9A32CD',
    'DarkOrchid4': '#68228B',
    'DarkRed': '#8B0000',
    'DarkSalmon': '#E9967A',
    'DarkSeaGreen': '#8FBC8F',
    'DarkSeaGreen1': '#C1FFC1',
    'DarkSeaGreen2': '#B4EEB4',
    'DarkSeaGreen3': '#9BCD9B',
    'DarkSeaGreen4': '#698B69',
    'DarkSlateBlue': '#483D8B',
    'DarkSlateGray': '#2F4F4F',
    'DarkSlateGray1': '#97FFFF',
    'DarkSlateGray2': '#8DEEEE',
    'DarkSlateGray3': '#79CDCD',
    'DarkSlateGray4': '#528B8B',
    'DarkSlateGrey': '#2F4F4F',
    'DarkTurquoise': '#00CED1',
    'DarkViolet': '#9400D3',
    'deep pink': '#FF1493',
    'deep sky blue': '#00BFFF',
    'DeepPink': '#FF1493',
    'DeepPink1': '#FF1493',
    'DeepPink2': '#EE1289',
    'DeepPink3': '#CD1076',
    'DeepPink4': '#8B0A50',
    'DeepSkyBlue': '#00BFFF',
    'DeepSkyBlue1': '#00BFFF',
    'DeepSkyBlue2': '#00B2EE',
    'DeepSkyBlue3': '#009ACD',
    'DeepSkyBlue4': '#00688B',
    'dim gray': '#696969',
    'dim grey': '#696969',
    'DimGray': '#696969',
    'DimGrey': '#696969',
    'dodger blue': '#1E90FF',
    'DodgerBlue': '#1E90FF',
    'DodgerBlue1': '#1E90FF',
    'DodgerBlue2': '#1C86EE',
    'DodgerBlue3': '#1874CD',
    'DodgerBlue4': '#104E8B',
    'firebrick': '#B22222',
    'firebrick1': '#FF3030',
    'firebrick2': '#EE2C2C',
    'firebrick3': '#CD2626',
    'firebrick4': '#8B1A1A',
    'floral white': '#FFFAF0',
    'FloralWhite': '#FFFAF0',
    'forest green': '#228B22',
    'ForestGreen': '#228B22',
    'gainsboro': '#DCDCDC',
    'ghost white': '#F8F8FF',
    'GhostWhite': '#F8F8FF',
    'gold': '#FFD700',
    'gold1': '#FFD700',
    'gold2': '#EEC900',
    'gold3': '#CDAD00',
    'gold4': '#8B7500',
    'goldenrod': '#DAA520',
    'goldenrod1': '#FFC125',
    'goldenrod2': '#EEB422',
    'goldenrod3': '#CD9B1D',
    'goldenrod4': '#8B6914',
    'gray': '#BEBEBE',
    'gray0': '#000000',
    'gray1': '#030303',
    'gray2': '#050505',
    'gray3': '#080808',
    'gray4': '#0A0A0A',
    'gray5': '#0D0D0D',
    'gray6': '#0F0F0F',
    'gray7': '#121212',
    'gray8': '#141414',
    'gray9': '#171717',
    'gray10': '#1A1A1A',
    'gray11': '#1C1C1C',
    'gray12': '#1F1F1F',
    'gray13': '#212121',
    'gray14': '#242424',
    'gray15': '#262626',
    'gray16': '#292929',
    'gray17': '#2B2B2B',
    'gray18': '#2E2E2E',
    'gray19': '#303030',
    'gray20': '#333333',
    'gray21': '#363636',
    'gray22': '#383838',
    'gray23': '#3B3B3B',
    'gray24': '#3D3D3D',
    'gray25': '#404040',
    'gray26': '#424242',
    'gray27': '#454545',
    'gray28': '#474747',
    'gray29': '#4A4A4A',
    'gray30': '#4D4D4D',
    'gray31': '#4F4F4F',
    'gray32': '#525252',
    'gray33': '#545454',
    'gray34': '#575757',
    'gray35': '#595959',
    'gray36': '#5C5C5C',
    'gray37': '#5E5E5E',
    'gray38': '#616161',
    'gray39': '#636363',
    'gray40': '#666666',
    'gray41': '#696969',
    'gray42': '#6B6B6B',
    'gray43': '#707070',
    'gray44': '#707070',
    'gray45': '#707070',
    'gray46': '#757575',
    'gray47': '#787878',
    'gray48': '#7A7A7A',
    '#707070': '#707070',
    '#707070': '#7F7F7F',
    'gray51': '#828282',
    'gray52': '#858585',
    'gray53': '#878787',
    'gray54': '#8A8A8A',
    '#707070': '#8C8C8C',
    'gray56': '#8F8F8F',
    'gray57': '#919191',
    'gray58': '#949494',
    'gray59': '#969696',
    'gray60': '#999999',
    'gray61': '#9C9C9C',
    'gray62': '#9E9E9E',
    'gray63': '#A1A1A1',
    'gray64': '#A3A3A3',
    'gray65': '#A6A6A6',
    'gray66': '#A8A8A8',
    'gray67': '#ABABAB',
    'gray68': '#ADADAD',
    'gray69': '#B0B0B0',
    'gray70': '#B3B3B3',
    'gray71': '#B5B5B5',
    'gray72': '#B8B8B8',
    'gray73': '#BABABA',
    'gray74': '#BDBDBD',
    'gray75': '#BFBFBF',
    'gray76': '#C2C2C2',
    'gray77': '#C4C4C4',
    'gray78': '#C7C7C7',
    'gray79': '#C9C9C9',
    'gray80': '#CCCCCC',
    'gray81': '#CFCFCF',
    'gray82': '#D1D1D1',
    'gray83': '#D4D4D4',
    'gray84': '#D6D6D6',
    'gray85': '#D9D9D9',
    'gray86': '#DBDBDB',
    'gray87': '#DEDEDE',
    'gray88': '#E0E0E0',
    'gray89': '#E3E3E3',
    'gray90': '#E5E5E5',
    'gray91': '#E8E8E8',
    'gray92': '#EBEBEB',
    'gray93': '#EDEDED',
    'gray94': '#F0F0F0',
    'gray95': '#F2F2F2',
    'gray96': '#F5F5F5',
    'gray97': '#F7F7F7',
    'gray98': '#FAFAFA',
    'gray99': '#FCFCFC',
    'gray100': '#FFFFFF',
    'green': '#00FF00',
    'green yellow': '#ADFF2F',
    'green1': '#00FF00',
    'green2': '#00EE00',
    'green3': '#00CD00',
    'green4': '#008B00',
    'GreenYellow': '#ADFF2F',
    'grey': '#BEBEBE',
    'grey0': '#000000',
    'grey1': '#030303',
    'grey2': '#050505',
    'grey3': '#080808',
    'grey4': '#0A0A0A',
    'grey5': '#0D0D0D',
    'grey6': '#0F0F0F',
    'grey7': '#121212',
    'grey8': '#141414',
    'grey9': '#171717',
    'grey10': '#1A1A1A',
    'grey11': '#1C1C1C',
    'grey12': '#1F1F1F',
    'grey13': '#212121',
    'grey14': '#242424',
    'grey15': '#262626',
    'grey16': '#292929',
    'grey17': '#2B2B2B',
    'grey18': '#2E2E2E',
    'grey19': '#303030',
    'grey20': '#333333',
    'grey21': '#363636',
    'grey22': '#383838',
    'grey23': '#3B3B3B',
    'grey24': '#3D3D3D',
    'grey25': '#404040',
    'grey26': '#424242',
    'grey27': '#454545',
    'grey28': '#474747',
    'grey29': '#4A4A4A',
    'grey30': '#4D4D4D',
    'grey31': '#4F4F4F',
    'grey32': '#525252',
    'grey33': '#545454',
    'grey34': '#575757',
    'grey35': '#595959',
    'grey36': '#5C5C5C',
    'grey37': '#5E5E5E',
    'grey38': '#616161',
    'grey39': '#636363',
    'grey40': '#666666',
    'grey41': '#696969',
    'grey42': '#6B6B6B',
    'grey43': '#707070',
    'grey44': '#707070',
    'grey45': '#707070',
    'grey46': '#757575',
    'grey47': '#787878',
    'grey48': '#7A7A7A',
    'grey49': '#707070',
    'grey50': '#7F7F7F',
    'grey51': '#828282',
    'grey52': '#858585',
    'grey53': '#878787',
    'grey54': '#8A8A8A',
    'grey55': '#8C8C8C',
    'grey56': '#8F8F8F',
    'grey57': '#919191',
    'grey58': '#949494',
    'grey59': '#969696',
    'grey60': '#999999',
    'grey61': '#9C9C9C',
    'grey62': '#9E9E9E',
    'grey63': '#A1A1A1',
    'grey64': '#A3A3A3',
    'grey65': '#A6A6A6',
    'grey66': '#A8A8A8',
    'grey67': '#ABABAB',
    'grey68': '#ADADAD',
    'grey69': '#B0B0B0',
    'grey70': '#B3B3B3',
    'grey71': '#B5B5B5',
    'grey72': '#B8B8B8',
    'grey73': '#BABABA',
    'grey74': '#BDBDBD',
    'grey75': '#BFBFBF',
    'grey76': '#C2C2C2',
    'grey77': '#C4C4C4',
    'grey78': '#C7C7C7',
    'grey79': '#C9C9C9',
    'grey80': '#CCCCCC',
    'grey81': '#CFCFCF',
    'grey82': '#D1D1D1',
    'grey83': '#D4D4D4',
    'grey84': '#D6D6D6',
    'grey85': '#D9D9D9',
    'grey86': '#DBDBDB',
    'grey87': '#DEDEDE',
    'grey88': '#E0E0E0',
    'grey89': '#E3E3E3',
    'grey90': '#E5E5E5',
    'grey91': '#E8E8E8',
    'grey92': '#EBEBEB',
    'grey93': '#EDEDED',
    'grey94': '#F0F0F0',
    'grey95': '#F2F2F2',
    'grey96': '#F5F5F5',
    'grey97': '#F7F7F7',
    'grey98': '#FAFAFA',
    'grey99': '#FCFCFC',
    'grey100': '#FFFFFF',
    'honeydew': '#F0FFF0',
    'honeydew1': '#F0FFF0',
    'honeydew2': '#E0EEE0',
    'honeydew3': '#C1CDC1',
    'honeydew4': '#838B83',
    'hot pink': '#FF69B4',
    'HotPink': '#FF69B4',
    'HotPink1': '#FF6EB4',
    'HotPink2': '#EE6AA7',
    'HotPink3': '#CD6090',
    'HotPink4': '#8B3A62',
    'indian red': '#CD5C5C',
    'IndianRed': '#CD5C5C',
    'IndianRed1': '#FF6A6A',
    'IndianRed2': '#EE6363',
    'IndianRed3': '#CD5555',
    'IndianRed4': '#8B3A3A',
    'ivory': '#FFFFF0',
    'ivory1': '#FFFFF0',
    'ivory2': '#EEEEE0',
    'ivory3': '#CDCDC1',
    'ivory4': '#8B8B83',
    'khaki': '#F0E68C',
    'khaki1': '#FFF68F',
    'khaki2': '#EEE685',
    'khaki3': '#CDC673',
    'khaki4': '#8B864E',
    'lavender': '#E6E6FA',
    'lavender blush': '#FFF0F5',
    'LavenderBlush': '#FFF0F5',
    'LavenderBlush1': '#FFF0F5',
    'LavenderBlush2': '#EEE0E5',
    'LavenderBlush3': '#CDC1C5',
    'LavenderBlush4': '#8B8386',
    'lawn green': '#7CFC00',
    'LawnGreen': '#7CFC00',
    'lemon chiffon': '#FFFACD',
    'LemonChiffon': '#FFFACD',
    'LemonChiffon1': '#FFFACD',
    'LemonChiffon2': '#EEE9BF',
    'LemonChiffon3': '#CDC9A5',
    'LemonChiffon4': '#8B8970',
    'light blue': '#ADD8E6',
    'light coral': '#F08080',
    'light cyan': '#E0FFFF',
    'light goldenrod': '#EEDD82',
    'light goldenrod yellow': '#FAFAD2',
    'light gray': '#D3D3D3',
    'light green': '#90EE90',
    'light grey': '#D3D3D3',
    'light pink': '#FFB6C1',
    'light salmon': '#FFA07A',
    'light sea green': '#20B2AA',
    'light sky blue': '#87CEFA',
    'light slate blue': '#8470FF',
    'light slate gray': '#778899',
    'light slate grey': '#778899',
    'light steel blue': '#B0C4DE',
    'light yellow': '#FFFFE0',
    'LightBlue': '#ADD8E6',
    'LightBlue1': '#BFEFFF',
    'LightBlue2': '#B2DFEE',
    'LightBlue3': '#9AC0CD',
    'LightBlue4': '#68838B',
    'LightCoral': '#F08080',
    'LightCyan': '#E0FFFF',
    'LightCyan1': '#E0FFFF',
    'LightCyan2': '#D1EEEE',
    'LightCyan3': '#B4CDCD',
    'LightCyan4': '#7A8B8B',
    'LightGoldenrod': '#EEDD82',
    'LightGoldenrod1': '#FFEC8B',
    'LightGoldenrod2': '#EEDC82',
    'LightGoldenrod3': '#CDBE70',
    'LightGoldenrod4': '#8B814C',
    'LightGoldenrodYellow': '#FAFAD2',
    'LightGray': '#D3D3D3',
    'LightGreen': '#90EE90',
    'LightGrey': '#D3D3D3',
    'LightPink': '#FFB6C1',
    'LightPink1': '#FFAEB9',
    'LightPink2': '#EEA2AD',
    'LightPink3': '#CD8C95',
    'LightPink4': '#8B5F65',
    'LightSalmon': '#FFA07A',
    'LightSalmon1': '#FFA07A',
    'LightSalmon2': '#EE9572',
    'LightSalmon3': '#CD8162',
    'LightSalmon4': '#8B5742',
    'LightSeaGreen': '#20B2AA',
    'LightSkyBlue': '#87CEFA',
    'LightSkyBlue1': '#B0E2FF',
    'LightSkyBlue2': '#A4D3EE',
    'LightSkyBlue3': '#8DB6CD',
    'LightSkyBlue4': '#607B8B',
    'LightSlateBlue': '#8470FF',
    'LightSlateGray': '#778899',
    'LightSlateGrey': '#778899',
    'LightSteelBlue': '#B0C4DE',
    'LightSteelBlue1': '#CAE1FF',
    'LightSteelBlue2': '#BCD2EE',
    'LightSteelBlue3': '#A2B5CD',
    'LightSteelBlue4': '#6E7B8B',
    'LightYellow': '#FFFFE0',
    'LightYellow1': '#FFFFE0',
    'LightYellow2': '#EEEED1',
    'LightYellow3': '#CDCDB4',
    'LightYellow4': '#8B8B7A',
    'lime green': '#32CD32',
    'LimeGreen': '#32CD32',
    'linen': '#FAF0E6',
    'magenta': '#FF00FF',
    'magenta1': '#FF00FF',
    'magenta2': '#EE00EE',
    'magenta3': '#CD00CD',
    'magenta4': '#8B008B',
    'maroon': '#B03060',
    'maroon1': '#FF34B3',
    'maroon2': '#EE30A7',
    'maroon3': '#CD2990',
    'maroon4': '#8B1C62',
    'medium aquamarine': '#66CDAA',
    'medium blue': '#0000CD',
    'medium orchid': '#BA55D3',
    'medium purple': '#9370DB',
    'medium sea green': '#3CB371',
    'medium slate blue': '#7B68EE',
    'medium spring green': '#00FA9A',
    'medium turquoise': '#48D1CC',
    'medium violet red': '#C71585',
    'MediumAquamarine': '#66CDAA',
    'MediumBlue': '#0000CD',
    'MediumOrchid': '#BA55D3',
    'MediumOrchid1': '#E066FF',
    'MediumOrchid2': '#D15FEE',
    'MediumOrchid3': '#B452CD',
    'MediumOrchid4': '#7A378B',
    'MediumPurple': '#9370DB',
    'MediumPurple1': '#AB82FF',
    'MediumPurple2': '#9F79EE',
    'MediumPurple3': '#8968CD',
    'MediumPurple4': '#5D478B',
    'MediumSeaGreen': '#3CB371',
    'MediumSlateBlue': '#7B68EE',
    'MediumSpringGreen': '#00FA9A',
    'MediumTurquoise': '#48D1CC',
    'MediumVioletRed': '#C71585',
    'midnight blue': '#191970',
    'MidnightBlue': '#191970',
    'mint cream': '#F5FFFA',
    'MintCream': '#F5FFFA',
    'misty rose': '#FFE4E1',
    'MistyRose': '#FFE4E1',
    'MistyRose1': '#FFE4E1',
    'MistyRose2': '#EED5D2',
    'MistyRose3': '#CDB7B5',
    'MistyRose4': '#8B7D7B',
    'moccasin': '#FFE4B5',
    'navajo white': '#FFDEAD',
    'NavajoWhite': '#FFDEAD',
    'NavajoWhite1': '#FFDEAD',
    'NavajoWhite2': '#EECFA1',
    'NavajoWhite3': '#CDB38B',
    'NavajoWhite4': '#8B795E',
    'navy': '#000080',
    'navy blue': '#000080',
    'NavyBlue': '#000080',
    'old lace': '#FDF5E6',
    'OldLace': '#FDF5E6',
    'olive drab': '#6B8E23',
    'OliveDrab': '#6B8E23',
    'OliveDrab1': '#C0FF3E',
    'OliveDrab2': '#B3EE3A',
    'OliveDrab3': '#9ACD32',
    'OliveDrab4': '#698B22',
    'orange': '#FFA500',
    'orange red': '#FF4500',
    'orange1': '#FFA500',
    'orange2': '#EE9A00',
    'orange3': '#CD8500',
    'orange4': '#8B5A00',
    'OrangeRed': '#FF4500',
    'OrangeRed1': '#FF4500',
    'OrangeRed2': '#EE4000',
    'OrangeRed3': '#CD3700',
    'OrangeRed4': '#8B2500',
    'orchid': '#DA70D6',
    'orchid1': '#FF83FA',
    'orchid2': '#EE7AE9',
    'orchid3': '#CD69C9',
    'orchid4': '#8B4789',
    'pale goldenrod': '#EEE8AA',
    'pale green': '#98FB98',
    'pale turquoise': '#AFEEEE',
    'pale violet red': '#DB7093',
    'PaleGoldenrod': '#EEE8AA',
    'PaleGreen': '#98FB98',
    'PaleGreen1': '#9AFF9A',
    'PaleGreen2': '#90EE90',
    'PaleGreen3': '#7CCD7C',
    'PaleGreen4': '#548B54',
    'PaleTurquoise': '#AFEEEE',
    'PaleTurquoise1': '#BBFFFF',
    'PaleTurquoise2': '#AEEEEE',
    'PaleTurquoise3': '#96CDCD',
    'PaleTurquoise4': '#668B8B',
    'PaleVioletRed': '#DB7093',
    'PaleVioletRed1': '#FF82AB',
    'PaleVioletRed2': '#EE799F',
    'PaleVioletRed3': '#CD687F',
    'PaleVioletRed4': '#8B475D',
    'papaya whip': '#FFEFD5',
    'PapayaWhip': '#FFEFD5',
    'peach puff': '#FFDAB9',
    'PeachPuff': '#FFDAB9',
    'PeachPuff1': '#FFDAB9',
    'PeachPuff2': '#EECBAD',
    'PeachPuff3': '#CDAF95',
    'PeachPuff4': '#8B7765',
    'peru': '#CD853F',
    'pink': '#FFC0CB',
    'pink1': '#FFB5C5',
    'pink2': '#EEA9B8',
    'pink3': '#CD919E',
    'pink4': '#8B636C',
    'plum': '#DDA0DD',
    'plum1': '#FFBBFF',
    'plum2': '#EEAEEE',
    'plum3': '#CD96CD',
    'plum4': '#8B668B',
    'powder blue': '#B0E0E6',
    'PowderBlue': '#B0E0E6',
    'purple': '#A020F0',
    'purple1': '#9B30FF',
    'purple2': '#912CEE',
    'purple3': '#7D26CD',
    'purple4': '#551A8B',
    'red': '#FF0000',
    'red1': '#FF0000',
    'red2': '#EE0000',
    'red3': '#CD0000',
    'red4': '#8B0000',
    'rosy brown': '#BC8F8F',
    'RosyBrown': '#BC8F8F',
    'RosyBrown1': '#FFC1C1',
    'RosyBrown2': '#EEB4B4',
    'RosyBrown3': '#CD9B9B',
    'RosyBrown4': '#8B6969',
    'royal blue': '#4169E1',
    'RoyalBlue': '#4169E1',
    'RoyalBlue1': '#4876FF',
    'RoyalBlue2': '#436EEE',
    'RoyalBlue3': '#3A5FCD',
    'RoyalBlue4': '#27408B',
    'saddle brown': '#8B4513',
    'SaddleBrown': '#8B4513',
    'salmon': '#FA8072',
    'salmon1': '#FF8C69',
    'salmon2': '#EE8262',
    'salmon3': '#CD7054',
    'salmon4': '#8B4C39',
    'sandy brown': '#F4A460',
    'SandyBrown': '#F4A460',
    'sea green': '#2E8B57',
    'SeaGreen': '#2E8B57',
    'SeaGreen1': '#54FF9F',
    'SeaGreen2': '#4EEE94',
    'SeaGreen3': '#43CD80',
    'SeaGreen4': '#2E8B57',
    'seashell': '#FFF5EE',
    'seashell1': '#FFF5EE',
    'seashell2': '#EEE5DE',
    'seashell3': '#CDC5BF',
    'seashell4': '#8B8682',
    'sienna': '#A0522D',
    'sienna1': '#FF8247',
    'sienna2': '#EE7942',
    'sienna3': '#CD6839',
    'sienna4': '#8B4726',
    'sky blue': '#87CEEB',
    'SkyBlue': '#87CEEB',
    'SkyBlue1': '#87CEFF',
    'SkyBlue2': '#7EC0EE',
    'SkyBlue3': '#6CA6CD',
    'SkyBlue4': '#4A708B',
    'slate blue': '#6A5ACD',
    'slate gray': '#708090',
    'slate grey': '#708090',
    'SlateBlue': '#6A5ACD',
    'SlateBlue1': '#836FFF',
    'SlateBlue2': '#7A67EE',
    'SlateBlue3': '#6959CD',
    'SlateBlue4': '#473C8B',
    'SlateGray': '#708090',
    'SlateGray1': '#C6E2FF',
    'SlateGray2': '#B9D3EE',
    'SlateGray3': '#9FB6CD',
    'SlateGray4': '#6C7B8B',
    'SlateGrey': '#708090',
    'snow': '#FFFAFA',
    'snow1': '#FFFAFA',
    'snow2': '#EEE9E9',
    'snow3': '#CDC9C9',
    'snow4': '#8B8989',
    'spring green': '#00FF7F',
    'SpringGreen': '#00FF7F',
    'SpringGreen1': '#00FF7F',
    'SpringGreen2': '#00EE76',
    'SpringGreen3': '#00CD66',
    'SpringGreen4': '#008B45',
    'steel blue': '#4682B4',
    'SteelBlue': '#4682B4',
    'SteelBlue1': '#63B8FF',
    'SteelBlue2': '#5CACEE',
    'SteelBlue3': '#4F94CD',
    'SteelBlue4': '#36648B',
    'tan': '#D2B48C',
    'tan1': '#FFA54F',
    'tan2': '#EE9A49',
    'tan3': '#CD853F',
    'tan4': '#8B5A2B',
    'thistle': '#D8BFD8',
    'thistle1': '#FFE1FF',
    'thistle2': '#EED2EE',
    'thistle3': '#CDB5CD',
    'thistle4': '#8B7B8B',
    'tomato': '#FF6347',
    'tomato1': '#FF6347',
    'tomato2': '#EE5C42',
    'tomato3': '#CD4F39',
    'tomato4': '#8B3626',
    'turquoise': '#40E0D0',
    'turquoise1': '#00F5FF',
    'turquoise2': '#00E5EE',
    'turquoise3': '#00C5CD',
    'turquoise4': '#00868B',
    'violet': '#EE82EE',
    'violet red': '#D02090',
    'VioletRed': '#D02090',
    'VioletRed1': '#FF3E96',
    'VioletRed2': '#EE3A8C',
    'VioletRed3': '#CD3278',
    'VioletRed4': '#8B2252',
    'wheat': '#F5DEB3',
    'wheat1': '#FFE7BA',
    'wheat2': '#EED8AE',
    'wheat3': '#CDBA96',
    'wheat4': '#8B7E66',
    'white': '#FFFFFF',
    'white smoke': '#F5F5F5',
    'WhiteSmoke': '#F5F5F5',
    'yellow': '#FFFF00',
    'yellow green': '#9ACD32',
    'yellow1': '#FFFF00',
    'yellow2': '#EEEE00',
    'yellow3': '#CDCD00',
    'yellow4': '#8B8B00',
    'YellowGreen': '#9ACD32',
}


class TextEditor_(tk.Text):
    def __init__(self,
                 parent,
                 font = ("TkHeadingFont", 10),
                 text = None,
                 state = "normal"):
        tk.Text.__init__(self,
                         parent,
                         font = font,
                         state = state,
                         spacing1 = 2,
                         spacing2 = 2,
                         undo = True)
        if text is not None:
            self.insert(1.0, text)
        self.rc_popup_menu = tk.Menu(self, tearoff = 0)
        self.rc_popup_menu.add_command(label = "Select all (Ctrl-a)",
                                       command = self.select_all)
        self.rc_popup_menu.add_separator()
        self.rc_popup_menu.add_command(label = "Cut (Ctrl-x)",
                                       command = self.cut)
        self.rc_popup_menu.add_separator()
        self.rc_popup_menu.add_command(label = "Copy (Ctrl-c)",
                                       command = self.copy)
        self.rc_popup_menu.add_separator()
        self.rc_popup_menu.add_command(label = "Paste (Ctrl-v)",
                                       command = self.paste)
        self.bind("<1>", lambda event: self.focus_set())
        if str(get_os()) == "Darwin":
            self.bind("<2>", self.rc)
        else:
            self.bind("<3>", self.rc)
        #self.bind("<Alt-Return>", self.add_newline)
        
    def add_newline(self, event):
        #self.insert("end", "\n")
        pass
    
    def rc(self,event):
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root, event.y_root)
        
    def select_all(self, event = None):
        self.event_generate("<Control-a>")
        return "break"
    
    def cut(self, event = None):
        self.event_generate("<Control-x>")
        return "break"
    
    def copy(self, event = None):
        self.event_generate("<Control-c>")
        return "break"
    
    def paste(self, event = None):
        self.event_generate("<Control-v>")
        return "break"


class TextEditor(tk.Frame):
    def __init__(self,
                 parent,
                 font = ("TkHeadingFont", 10),
                 text = None,
                 state = "normal",
                 width = None,
                 height = None):
        tk.Frame.__init__(self,
                          parent,
                          height = height,
                          width = width)
        self.textedit = TextEditor_(self,
                                    font = font,
                                    text = text,
                                    state = state)
        self.textedit.grid(row = 0,
                           column = 0,
                           sticky = "nswe")
        self.grid_columnconfigure(0, weight = 1)
        self.grid_rowconfigure(0, weight = 1)
        self.grid_propagate(False)
        self.textedit.focus_set()
        
    def get(self):
        return self.textedit.get("1.0", "end").rstrip()


class TopLeftRectangle(tk.Canvas):
    def __init__(self,
                 parentframe = None,
                 main_canvas = None,
                 row_index_canvas = None,
                 header_canvas = None,
                 background = None,
                 foreground = None):
        tk.Canvas.__init__(self,
                           parentframe,
                           background = background,
                           highlightthickness = 0)
        self.parentframe = parentframe
        self.rectangle_foreground = foreground
        self.MT = main_canvas
        self.RI = row_index_canvas
        self.CH = header_canvas
        self.config(width = self.RI.current_width, height = self.CH.current_height)
        self.extra_motion_func = None
        self.extra_b1_press_func = None
        self.extra_b1_motion_func = None
        self.extra_b1_release_func = None
        self.extra_double_b1_func = None
        self.MT.TL = self
        self.RI.TL = self
        self.CH.TL = self
        w = self.RI.current_width - 1
        h = self.CH.current_height - 1
        half_x = floor(w / 2)
        self.create_rectangle(1, 1, half_x, h, fill = self.rectangle_foreground, outline = "", tag = "rw")
        self.create_rectangle(half_x + 1, 1, w, h, fill = self.rectangle_foreground, outline = "", tag = "rh")
        self.bind("<Motion>", self.mouse_motion)
        self.bind("<ButtonPress-1>", self.b1_press)
        self.bind("<B1-Motion>", self.b1_motion)
        self.bind("<ButtonRelease-1>", self.b1_release)
        self.bind("<Double-Button-1>", self.double_b1)

    def basic_bindings(self, onoff = "enable"):
        if onoff == "enable":
            self.bind("<Motion>", self.mouse_motion)
            self.bind("<ButtonPress-1>", self.b1_press)
            self.bind("<B1-Motion>", self.b1_motion)
            self.bind("<ButtonRelease-1>", self.b1_release)
            self.bind("<Double-Button-1>", self.double_b1)
        elif onoff == "disable":
            self.unbind("<Motion>")
            self.unbind("<ButtonPress-1>")
            self.unbind("<B1-Motion>")
            self.unbind("<ButtonRelease-1>")
            self.unbind("<Double-Button-1>")

    def set_dimensions(self, new_w = None, new_h = None):
        if new_w:
            self.config(width = new_w)
            w = new_w - 1
            h = self.winfo_height() - 1
        if new_h:
            self.config(height = new_h)
            w = self.winfo_width() - 1
            h = new_h - 1
        half_x = floor(w / 2)
        self.coords("rw", 1, 1, half_x, h)
        self.coords("rh", half_x + 1, 1, w, h)

    def mouse_motion(self, event = None):
        self.MT.reset_mouse_motion_creations()
        if self.extra_motion_func is not None:
            self.extra_motion_func()

    def b1_press(self, event = None):
        self.focus_set()
        rect = self.find_closest(event.x, event.y)
        if rect[0] % 2:
            if self.RI.width_resizing_enabled:
                self.RI.set_width(self.RI.default_width, set_TL = True) # DEFAULT SIZE PARAMTER INSTEAD ????
        else:
            if self.CH.height_resizing_enabled:
                self.CH.set_height(self.MT.hdr_min_rh, set_TL = True)
        self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.extra_b1_press_func is not None:
            self.extra_b1_press_func(event)

    def b1_motion(self, event = None):
        if self.extra_b1_motion_func is not None:
            self.extra_b1_motion_func(event)

    def b1_release(self, event = None):
        if self.extra_b1_release_func is not None:
            self.extra_b1_release_func(event)

    def double_b1(self, event = None):
        self.focus_set()
        if self.extra_double_b1_func is not None:
            self.extra_double_b1_func(event)


class ColumnHeaders(tk.Canvas):
    def __init__(self,
                 parentframe = None,
                 main_canvas = None,
                 row_index_canvas = None,
                 max_colwidth = None,
                 max_header_height = None,
                 header_align = None,
                 header_background = None,
                 header_border_color = None,
                 header_grid_color = None,
                 header_foreground = None,
                 header_select_background = None,
                 header_select_foreground = None,
                 drag_and_drop_color = None,
                 resizing_line_color = None):
        tk.Canvas.__init__(self,parentframe,
                           background = header_background,
                           highlightthickness = 0)
        self.parentframe = parentframe
        self.extra_motion_func = None
        self.extra_b1_press_func = None
        self.extra_b1_motion_func = None
        self.extra_b1_release_func = None
        self.extra_double_b1_func = None
        self.ch_extra_drag_drop_func = None
        self.selection_binding_func = None
        self.drag_selection_binding_func = None
        self.max_colwidth = float(max_colwidth)
        self.max_header_height = float(max_header_height)
        self.current_height = None    # is set from within MainTable() __init__ or from Sheet parameters
        self.MT = main_canvas         # is set from within MainTable() __init__
        self.RI = row_index_canvas    # is set from within MainTable() __init__
        self.TL = None                # is set from within TopLeftRectangle() __init__
        self.text_color = header_foreground
        self.grid_color = header_grid_color
        self.header_border_color = header_border_color
        self.selected_cells_background = header_select_background
        self.selected_cells_foreground = header_select_foreground
        self.drag_and_drop_color = drag_and_drop_color
        self.resizing_line_color = resizing_line_color
        self.align = header_align
        self.width_resizing_enabled = False
        self.height_resizing_enabled = False
        self.double_click_resizing_enabled = False
        self.col_selection_enabled = False
        self.drag_and_drop_enabled = False
        self.dragged_col = None
        self.visible_col_dividers = []
        self.col_height_resize_bbox = tuple()
        self.selected_cells = defaultdict(int)
        self.highlighted_cells = {}
        self.rsz_w = None
        self.rsz_h = None
        self.new_col_height = 0
        self.currently_resizing_width = False
        self.currently_resizing_height = False
        self.bind("<Motion>", self.mouse_motion)
        self.bind("<ButtonPress-1>", self.b1_press)
        self.bind("<Shift-ButtonPress-1>",self.shift_b1_press)
        self.bind("<B1-Motion>", self.b1_motion)
        self.bind("<ButtonRelease-1>", self.b1_release)
        self.bind("<Double-Button-1>", self.double_b1)

    def basic_bindings(self, onoff = "enable"):
        if onoff == "enable":
            self.bind("<Motion>", self.mouse_motion)
            self.bind("<ButtonPress-1>", self.b1_press)
            self.bind("<B1-Motion>", self.b1_motion)
            self.bind("<ButtonRelease-1>", self.b1_release)
            self.bind("<Double-Button-1>", self.double_b1)
        elif onoff == "disable":
            self.unbind("<Motion>")
            self.unbind("<ButtonPress-1>")
            self.unbind("<B1-Motion>")
            self.unbind("<ButtonRelease-1>")
            self.unbind("<Double-Button-1>")

    def set_height(self, new_height,set_TL = False):
        self.current_height = new_height
        self.config(height = new_height)
        if set_TL:
            self.TL.set_dimensions(new_h = new_height)

    def enable_bindings(self, binding):
        if binding == "column_width_resize":
            self.width_resizing_enabled = True
        if binding == "column_height_resize":
            self.height_resizing_enabled = True
        if binding == "double_click_column_resize":
            self.double_click_resizing_enabled = True
        if binding == "column_select":
            self.col_selection_enabled = True
        if binding == "drag_and_drop":
            self.drag_and_drop_enabled = True

    def disable_bindings(self, binding):
        if binding == "column_width_resize":
            self.width_resizing_enabled = False
        if binding == "column_height_resize":
            self.height_resizing_enabled = False
        if binding == "double_click_column_resize":
            self.double_click_resizing_enabled = False
        if binding == "column_select":
            self.col_selection_enabled = False
        if binding == "drag_and_drop":
            self.drag_and_drop_enabled = False

    def check_mouse_position_width_resizers(self, event):
        x = self.canvasx(event.x)
        y = self.canvasy(event.y)
        ov = None
        for x1, y1, x2, y2 in self.visible_col_dividers:
            if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                ov = self.find_overlapping(x1, y1, x2, y2)
                break
        return ov

    def shift_b1_press(self, event):
        x = event.x
        c = self.MT.identify_col(x = x)
        if self.drag_and_drop_enabled or self.col_selection_enabled and self.rsz_h is None and self.rsz_w is None:
            if c < len(self.MT.col_positions) - 1:
                if c not in self.MT.selected_cols and self.col_selection_enabled:
                    c = int(c)
                    if self.MT.currently_selected and self.MT.currently_selected[0] == "column":
                        min_c = int(self.MT.currently_selected[1])
                        self.selected_cells = defaultdict(int)
                        self.RI.selected_cells = defaultdict(int)
                        self.MT.selected_cols = set()
                        self.MT.selected_rows = set()
                        self.MT.selected_cells = set()
                        if c > min_c:
                            for i in range(min_c, c + 1):
                                self.selected_cells[i] += 1
                                self.MT.selected_cols.add(i)
                        elif c < min_c:
                            for i in range(c, min_c + 1):
                                self.selected_cells[i] += 1
                                self.MT.selected_cols.add(i)
                    else:
                        self.select_col(c)
                    self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
                    if self.selection_binding_func is not None:
                        self.selection_binding_func(("column", c))
                elif c in self.MT.selected_cols:
                    self.dragged_col = c

    def mouse_motion(self, event):
        if not self.currently_resizing_height and not self.currently_resizing_width:
            x = self.canvasx(event.x)
            y = self.canvasy(event.y)
            mouse_over_resize = False
            if self.width_resizing_enabled and not mouse_over_resize:
                ov = self.check_mouse_position_width_resizers(event)
                if ov is not None:
                    for itm in ov:
                        tgs = self.gettags(itm)
                        if "v" == tgs[0]:
                            break
                    c = int(tgs[1])
                    self.rsz_w = c
                    self.config(cursor = "sb_h_double_arrow")
                    mouse_over_resize = True
                else:
                    self.rsz_w = None
            if self.height_resizing_enabled and not mouse_over_resize:
                try:
                    x1, y1, x2, y2 = self.col_height_resize_bbox[0], self.col_height_resize_bbox[1], self.col_height_resize_bbox[2], self.col_height_resize_bbox[3]
                    if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                        self.config(cursor = "sb_v_double_arrow")
                        self.rsz_h = True
                        mouse_over_resize = True
                    else:
                        self.rsz_h = None
                except:
                    self.rsz_h = None
            if not mouse_over_resize:
                self.MT.reset_mouse_motion_creations()
        if self.extra_motion_func is not None:
            self.extra_motion_func(event)
        
    def b1_press(self, event = None):
        self.focus_set()
        self.MT.unbind("<MouseWheel>")
        x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
        if self.check_mouse_position_width_resizers(event) is None:
            self.rsz_w = None
        if self.width_resizing_enabled and self.rsz_w is not None:
            self.currently_resizing_width = True
            x = self.MT.col_positions[self.rsz_w]
            line2x = self.MT.col_positions[self.rsz_w - 1]
            self.create_line(x, 0, x, self.current_height, width = 1, fill = self.resizing_line_color, tag = "rwl")
            self.MT.create_line(x, y1, x, y2, width = 1, fill = self.resizing_line_color, tag = "rwl")
            self.create_line(line2x, 0, line2x, self.current_height,width = 1, fill = self.resizing_line_color, tag = "rwl2")
            self.MT.create_line(line2x, y1, line2x, y2, width = 1, fill = self.resizing_line_color, tag = "rwl2")
        elif self.height_resizing_enabled and self.rsz_w is None and self.rsz_h is not None:
            self.currently_resizing_height = True
            y = event.y
            if y < self.MT.hdr_min_rh:
                y = int(self.MT.hdr_min_rh)
            self.new_col_height = y
            self.create_line(x1, y, x2, y, width = 1, fill = self.resizing_line_color, tag = "rhl")
        elif self.MT.identify_col(x = event.x, allow_end = False) is None:
            self.MT.deselect("all")
        elif self.col_selection_enabled and self.rsz_w is None and self.rsz_h is None:
            c = self.MT.identify_col(x = event.x)
            if c < len(self.MT.col_positions) - 1:
                self.select_col(c, redraw = True)
        if self.extra_b1_press_func is not None:
            self.extra_b1_press_func(event)
    
    def b1_motion(self, event):
        x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
        if self.width_resizing_enabled and self.rsz_w is not None and self.currently_resizing_width:
            x = self.canvasx(event.x)
            size = x - self.MT.col_positions[self.rsz_w - 1]
            if not size <= self.MT.min_cw and size < self.max_colwidth:
                self.delete("rwl")
                self.MT.delete("rwl")
                self.create_line(x, 0, x, self.current_height, width = 1, fill = self.resizing_line_color, tag = "rwl")
                self.MT.create_line(x, y1, x, y2, width = 1, fill = self.resizing_line_color, tag = "rwl")
        elif self.height_resizing_enabled and self.rsz_h is not None and self.currently_resizing_height:
            evy = event.y
            self.delete("rhl")
            self.MT.delete("rhl")
            if evy > self.current_height:
                y = self.MT.canvasy(evy - self.current_height)
                if evy > self.max_header_height:
                    evy = int(self.max_header_height)
                    y = self.MT.canvasy(evy - self.current_height)
                self.new_col_height = evy
                self.MT.create_line(x1, y, x2, y, width = 1, fill = self.resizing_line_color, tag = "rhl")
            else:
                y = evy
                if y < self.MT.hdr_min_rh:
                    y = int(self.MT.hdr_min_rh)
                self.new_col_height = y
                self.create_line(x1, y, x2, y, width = 1, fill = self.resizing_line_color, tag = "rhl")
        elif self.drag_and_drop_enabled and self.col_selection_enabled and self.MT.selected_cols and self.rsz_h is None and self.rsz_w is None and self.dragged_col is not None:
            x = self.canvasx(event.x)
            if x > 0 and x < self.MT.col_positions[-1]:
                x = event.x
                wend = self.winfo_width() 
                if x >= wend - 0:
                    if x >= wend + 15:
                        self.MT.xview_scroll(2, "units")
                        self.xview_scroll(2, "units")
                    else:
                        self.MT.xview_scroll(1, "units")
                        self.xview_scroll(1, "units")
                    self.MT.main_table_redraw_grid_and_text(redraw_header = True)
                elif x <= 0:
                    if x >= -40:
                        self.MT.xview_scroll(-1, "units")
                        self.xview_scroll(-1, "units")
                    else:
                        self.MT.xview_scroll(-2, "units")
                        self.xview_scroll(-2, "units")
                    self.MT.main_table_redraw_grid_and_text(redraw_header = True)
                rectw = self.MT.col_positions[max(self.MT.selected_cols) + 1] - self.MT.col_positions[min(self.MT.selected_cols)]
                start = self.canvasx(event.x - int(rectw / 2))
                end = self.canvasx(event.x + int(rectw / 2))
                self.delete("dd")
                self.create_rectangle(start, 0, end, self.current_height - 1, fill = self.drag_and_drop_color, outline = self.grid_color, tag = "dd")
                self.tag_raise("dd")
                self.tag_raise("t")
                self.tag_raise("v")
        elif self.MT.drag_selection_enabled and self.col_selection_enabled and self.rsz_h is None and self.rsz_w is None:
            end_col = self.MT.identify_col(x = event.x)
            if end_col < len(self.MT.col_positions) - 1 and len(self.MT.currently_selected) == 2:
                if self.MT.currently_selected[0] == "column":
                    start_col = self.MT.currently_selected[1]
                    self.MT.selected_cols = set()
                    self.MT.selected_rows = set()
                    self.MT.selected_cells = set()
                    self.RI.selected_cells = defaultdict(int)
                    self.selected_cells = defaultdict(int)
                    if end_col >= start_col:
                        for c in range(start_col, end_col + 1):
                            self.selected_cells[c] += 1
                            self.MT.selected_cols.add(c)
                    elif end_col < start_col:
                        for c in range(end_col, start_col + 1):
                            self.selected_cells[c] += 1
                            self.MT.selected_cols.add(c)
                                
                    if self.drag_selection_binding_func is not None:
                        self.drag_selection_binding_func(("columns", sorted([start_col, end_col])))
                if event.x > self.winfo_width():
                    try:
                        self.MT.xview_scroll(1, "units")
                        self.xview_scroll(1, "units")
                    except:
                        pass
                elif event.x < 0 and self.canvasx(self.winfo_width()) > 0:
                    try:
                        self.xview_scroll(-1, "units")
                        self.MT.xview_scroll(-1, "units")
                    except:
                        pass
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = False)
        if self.extra_b1_motion_func is not None:
            self.extra_b1_motion_func(event)
            
    def b1_release(self, event = None):
        self.MT.bind("<MouseWheel>", self.MT.mousewheel)
        if self.width_resizing_enabled and self.rsz_w is not None and self.currently_resizing_width:
            self.currently_resizing_width = False
            new_col_pos = self.coords("rwl")[0]
            self.delete("rwl", "rwl2")
            self.MT.delete("rwl", "rwl2")
            increment = new_col_pos - self.MT.col_positions[self.rsz_w]
            self.MT.col_positions[self.rsz_w + 1:] = [e + increment for e in islice(self.MT.col_positions, self.rsz_w + 1, len(self.MT.col_positions))]
            self.MT.col_positions[self.rsz_w] = new_col_pos
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        elif self.height_resizing_enabled and self.rsz_h is not None and self.currently_resizing_height:
            self.currently_resizing_height = False
            self.delete("rhl")
            self.MT.delete("rhl")
            self.set_height(self.new_col_height,set_TL = True)
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.drag_and_drop_enabled and self.col_selection_enabled and self.MT.selected_cols and self.rsz_h is None and self.rsz_w is None and self.dragged_col is not None:
            self.delete("dd")
            x = event.x
            c = self.MT.identify_col(x = x)
            if c != self.dragged_col and c is not None and c not in self.MT.selected_cols and len(self.MT.selected_cols) != (len(self.MT.col_positions) - 1):
                colsiter = list(self.MT.selected_cols)
                colsiter.sort()
                stins = colsiter[0]
                endins = colsiter[-1] + 1
                if self.dragged_col < c and c >= len(self.MT.col_positions) - 1:
                    c -= 1
                c_ = int(c)
                if c >= endins:
                    c += 1
                if self.ch_extra_drag_drop_func is not None:
                    self.ch_extra_drag_drop_func(self.MT.selected_cols, int(c_))
                else:
                    if self.MT.all_columns_displayed:
                        if stins > c:
                            for rn in range(len(self.MT.data_ref)):
                                self.MT.data_ref[rn][c:c] = self.MT.data_ref[rn][stins:endins]
                                self.MT.data_ref[rn][stins + len(colsiter):endins + len(colsiter)] = []
                            if not isinstance(self.MT.my_hdrs, int) and self.MT.my_hdrs:
                                try:
                                    self.MT.my_hdrs[c:c] = self.MT.my_hdrs[stins:endins]
                                    self.MT.my_hdrs[stins + len(colsiter):endins + len(colsiter)] = []
                                except:
                                    pass
                        else:
                            for rn in range(len(self.MT.data_ref)):
                                self.MT.data_ref[rn][c:c] = self.MT.data_ref[rn][stins:endins]
                                self.MT.data_ref[rn][stins:endins] = []
                            if not isinstance(self.MT.my_hdrs, int) and self.MT.my_hdrs:
                                try:
                                    self.MT.my_hdrs[c:c] = self.MT.my_hdrs[stins:endins]
                                    self.MT.my_hdrs[stins:endins] = []
                                except:
                                    pass
                    else:
                        c_ = int(c)
                        if c >= endins:
                            c += 1
                        if stins > c:
                            self.MT.displayed_columns[c:c] = self.MT.displayed_columns[stins:endins]
                            self.MT.displayed_columns[stins + len(colsiter):endins + len(colsiter)] = []
                            if not isinstance(self.MT.my_hdrs, int) and self.MT.my_hdrs:
                                try:
                                    self.MT.my_hdrs[c:c] = self.MT.my_hdrs[stins:endins]
                                    self.MT.my_hdrs[stins + len(colsiter):endins + len(colsiter)] = []
                                except:
                                    pass
                        else:
                            self.MT.displayed_columns[c:c] = self.MT.displayed_columns[stins:endins]
                            self.MT.displayed_columns[stins + len(colsiter):endins + len(colsiter)] = []
                            if not isinstance(self.MT.my_hdrs, int) and self.MT.my_hdrs:
                                try:
                                    self.MT.my_hdrs[c:c] = self.MT.my_hdrs[stins:endins]
                                    self.MT.my_hdrs[stins:endins] = []
                                except:
                                    pass
                cws = self.MT.parentframe.get_column_widths()
                if stins > c:
                    cws[c:c] = cws[stins:endins]
                    cws[stins + len(colsiter):endins + len(colsiter)] = []
                else:
                    cws[c:c] = cws[stins:endins]
                    cws[stins:endins] = []
                self.MT.parentframe.set_column_widths(cws)
                if (c_ - 1) + len(colsiter) > len(self.MT.col_positions) - 1:
                    sels_start = len(self.MT.col_positions) - 1 - len(colsiter)
                    newcolidxs = tuple(range(sels_start, len(self.MT.col_positions) - 1))
                else:
                    if c_ > endins:
                        c_ += 1
                        sels_start = c_ - len(colsiter)
                    else:
                        if c_ == endins and len(colsiter) == 1:
                            pass
                        else:
                            if c_ > endins:
                                c_ += 1
                            if c_ == endins:
                                c_ -= 1
                            if c_ < 0:
                                c_ = 0
                        sels_start = c_
                    newcolidxs = tuple(range(sels_start, sels_start + len(colsiter)))
                self.MT.selected_rows = set()
                self.MT.selected_cells = set()
                self.selected_cells = defaultdict(int)
                self.RI.selected_cells = defaultdict(int)
                self.MT.selected_cols = set()
                for colsel in newcolidxs:
                    self.MT.selected_cols.add(colsel)
                    self.selected_cells[colsel] += 1
                self.MT.undo_storage = deque(maxlen = 20)
                self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        self.dragged_col = None
        self.currently_resizing_width = False
        self.currently_resizing_height = False
        self.rsz_w = None
        self.rsz_h = None
        self.mouse_motion(event)
        if self.extra_b1_release_func is not None:
            self.extra_b1_release_func(event)

    def double_b1(self, event = None):
        self.focus_set()
        if self.double_click_resizing_enabled and self.width_resizing_enabled and self.rsz_w is not None and not self.currently_resizing_width:
            # condition check if trying to resize width:
            col = self.rsz_w - 1
            self.set_col_width(col)
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
            self.mouse_motion(event)
        self.rsz_w = None
        if self.extra_double_b1_func is not None:
            self.extra_double_b1_func(event)

    def select_col(self, c, redraw = False):
        c = int(c)
        self.selected_cells = defaultdict(int)
        self.selected_cells[c] += 1
        self.RI.selected_cells = defaultdict(int)
        self.MT.selected_cols = {c}
        self.MT.selected_rows = set()
        self.MT.selected_cells = set()
        self.MT.currently_selected = ("column", c)
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.selection_binding_func is not None:
            self.selection_binding_func(("column", c))

    def highlight_cells(self, c = 0, cells = tuple(), bg = None, fg = None, redraw = False):
        if bg is None and fg is None:
            return
        if cells:
            self.highlighted_cells = {c_: (bg, fg)  for c_ in cells}
        else:
            self.highlighted_cells[c] = (bg, fg)
        if redraw:
            self.MT.main_table_redraw_grid_and_text(True, False)

    def add_selection(self, c, redraw = False, run_binding_func = True):
        c = int(c)
        self.MT.currently_selected = ("column", c)
        self.selected_cells[c] += 1
        self.MT.selected_cols.add(c)
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.selection_binding_func is not None and run_binding_func:
            self.selection_binding_func(("column", c))

    def set_col_width(self, col, width = None, only_set_if_too_small = False):
        if col < 0:
            return
        if width is None:
            if self.MT.all_columns_displayed:
                try:
                    hw = self.MT.GetHdrTextWidth(self.GetLargestWidth(self.MT.my_hdrs[col])) + 10
                except:
                    hw = self.MT.GetHdrTextWidth(str(col)) + 10
                x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
                start_row, end_row = self.MT.get_visible_rows(y1, y2)
                dtw = 0
                for r in islice(self.MT.data_ref, start_row, end_row):
                    try:
                        w = self.MT.GetTextWidth(self.GetLargestWidth(r[col]))
                        if w > dtw:
                            dtw = w
                    except:
                        pass
            else:
                try:
                    hw = self.MT.GetHdrTextWidth(self.GetLargestWidth(self.MT.my_hdrs[self.MT.displayed_columns[col]])) + 10 
                except:
                    hw = self.MT.GetHdrTextWidth(str(col)) + 10
                x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
                start_row,end_row = self.MT.get_visible_rows(y1, y2)
                dtw = 0
                for r in islice(self.MT.data_ref, start_row,end_row):
                    try:
                        w = self.MT.GetTextWidth(self.GetLargestWidth(r[self.MT.displayed_columns[col]]))
                        if w > dtw:
                            dtw = w
                    except:
                        pass 
            dtw += 10
            if dtw > hw:
                width = dtw
            else:
                width = hw
        if width <= self.MT.min_cw:
            width = int(self.MT.min_cw)
        elif width > self.max_colwidth:
            width = int(self.max_colwidth)
        if only_set_if_too_small:
            if width <= self.MT.col_positions[col + 1] - self.MT.col_positions[col]:
                return
        new_col_pos = self.MT.col_positions[col] + width
        increment = new_col_pos - self.MT.col_positions[col + 1]
        self.MT.col_positions[col + 2:] = [e + increment for e in islice(self.MT.col_positions, col + 2, len(self.MT.col_positions))]
        self.MT.col_positions[col + 1] = new_col_pos

    def GetLargestWidth(self, cell):
        return max(cell.split("\n"), key = self.MT.GetTextWidth)

    def redraw_grid_and_text(self, last_col_line_pos, x1, x_stop, start_col, end_col):
        try:
            self.configure(scrollregion = (0, 0, last_col_line_pos + 150, self.current_height))
            self.delete("h", "v", "t", "s", "fv")
            self.visible_col_dividers = []
            x = self.MT.col_positions[start_col]
            self.create_line(x, 0, x, self.current_height, fill = self.grid_color, width = 1, tag = "fv")
            self.col_height_resize_bbox = (x1, self.current_height - 4, x_stop, self.current_height)
            yend = self.current_height - 5
            if self.width_resizing_enabled:
                for c in range(start_col + 1, end_col):
                    x = self.MT.col_positions[c]
                    self.visible_col_dividers.append((x - 4, 1, x + 4, yend))
                    self.create_line(x, 0, x, self.current_height, fill = self.grid_color, width = 1, tag = ("v", f"{c}"))
            else:
                for c in range(start_col + 1, end_col):
                    x = self.MT.col_positions[c]
                    self.create_line(x, 0, x, self.current_height, fill = self.grid_color, width = 1, tag = ("v", f"{c}"))
            top = self.canvasy(0)
            if self.MT.hdr_fl_ins + self.MT.hdr_half_txt_h > top:
                incfl = True
            else:
                incfl = False
            c_2 = self.selected_cells_background if self.selected_cells_background.startswith("#") else Color_Map_[self.selected_cells_background]
            if self.MT.all_columns_displayed:
                if self.align == "center":
                    for c in range(start_col, end_col - 1):
                        fc = self.MT.col_positions[c]
                        sc = self.MT.col_positions[c + 1]
                        if c in self.highlighted_cells and (c in self.selected_cells or c in self.MT.selected_cols):
                            c_1 = self.highlighted_cells[c][0] if self.highlighted_cells[c][0].startswith("#") else Color_Map_[self.highlighted_cells[c][0]]
                            self.create_rectangle(fc + 1,
                                                  0,
                                                  sc,
                                                  self.current_height - 1,
                                                  fill = (f"#{int((int(c_1[1:3], 16) + int(c_2[1:3], 16)) / 2):02X}" +
                                                          f"{int((int(c_1[3:5], 16) + int(c_2[3:5], 16)) / 2):02X}" +
                                                          f"{int((int(c_1[5:], 16) + int(c_2[5:], 16)) / 2):02X}"),
                                                  outline = "",
                                                  tag = "s")
                            tf = self.selected_cells_foreground if self.highlighted_cells[c][1] is None else self.highlighted_cells[c][1]
                        elif c in (self.MT.selected_cols or self.selected_cells):
                            self.create_rectangle(fc + 1, 0, sc, self.current_height - 1, fill = self.selected_cells_background, outline = "", tag = "s")
                            tf = self.selected_cells_foreground
                        elif c in self.highlighted_cells:
                            self.create_rectangle(fc + 1, 0, sc, self.current_height - 1, fill = self.highlighted_cells[c][0], outline = "", tag = "s")
                            tf = self.text_color if self.highlighted_cells[c][1] is None else self.highlighted_cells[c][1]
                        else:
                            tf = self.text_color
                        if fc + 7 > x_stop:
                            continue
                        mw = sc - fc - 5
                        x = fc + floor(mw / 2)
                        if isinstance(self.MT.my_hdrs, int):
                            try:
                                lns = self.MT.data_ref[self.MT.my_hdrs][c].split("\n")
                            except:
                                lns = (f"{c + 1}", )
                        else:
                            try:
                                lns = self.MT.my_hdrs[c].split("\n")
                            except:
                                lns = (f"{c + 1}", )
                        y = self.MT.hdr_fl_ins
                        if incfl:
                            fl = lns[0]
                            t = self.create_text(x, y, text = fl, fill = tf, font = self.MT.my_hdr_font, anchor = "center", tag = "t")
                            wd = self.bbox(t)
                            wd = wd[2] - wd[0]
                            if wd > mw:
                                tl = len(fl)
                                slce = tl - floor(tl * (mw / wd))
                                if slce % 2:
                                    slce += 1
                                else:
                                    slce += 2
                                slce = int(slce / 2)
                                fl = fl[slce:tl - slce]
                                self.itemconfig(t, text = fl)
                                wd = self.bbox(t)
                                while wd[2] - wd[0] > mw:
                                    fl = fl[1: - 1]
                                    self.itemconfig(t, text = fl)
                                    wd = self.bbox(t)
                        if len(lns) > 1:
                            stl = int((top - y) / self.MT.hdr_xtra_lines_increment) - 1
                            if stl < 1:
                                stl = 1
                            y += (stl * self.MT.hdr_xtra_lines_increment)
                            if y + self.MT.hdr_half_txt_h < self.current_height:
                                for i in range(stl, len(lns)):
                                    txt = lns[i]
                                    t = self.create_text(x, y, text = txt, fill = tf, font = self.MT.my_hdr_font, anchor = "center", tag = "t")
                                    wd = self.bbox(t)
                                    wd = wd[2] - wd[0]
                                    if wd > mw:
                                        tl = len(txt)
                                        slce = tl - floor(tl * (mw / wd))
                                        if slce % 2:
                                            slce += 1
                                        else:
                                            slce += 2
                                        slce = int(slce / 2)
                                        txt = txt[slce:tl - slce]
                                        self.itemconfig(t, text = txt)
                                        wd = self.bbox(t)
                                        while wd[2] - wd[0] > mw:
                                            txt = txt[1: - 1]
                                            self.itemconfig(t, text = txt)
                                            wd = self.bbox(t)
                                    y += self.MT.hdr_xtra_lines_increment
                                    if y + self.MT.hdr_half_txt_h > self.current_height:
                                        break
                elif self.align == "w":
                    for c in range(start_col, end_col - 1):
                        fc = self.MT.col_positions[c]
                        sc = self.MT.col_positions[c + 1]
                        if c in self.highlighted_cells and (c in self.selected_cells or c in self.MT.selected_cols):
                            c_1 = self.highlighted_cells[c][0] if self.highlighted_cells[c][0].startswith("#") else Color_Map_[self.highlighted_cells[c][0]]
                            self.create_rectangle(fc + 1,
                                                  0,
                                                  sc,
                                                  self.current_height - 1,
                                                  fill = (f"#{int((int(c_1[1:3], 16) + int(c_2[1:3], 16)) / 2):02X}" +
                                                          f"{int((int(c_1[3:5], 16) + int(c_2[3:5], 16)) / 2):02X}" +
                                                          f"{int((int(c_1[5:], 16) + int(c_2[5:], 16)) / 2):02X}"),
                                                  outline = "",
                                                  tag = "s")
                            tf = self.selected_cells_foreground if self.highlighted_cells[c][1] is None else self.highlighted_cells[c][1]
                        elif c in (self.MT.selected_cols or self.selected_cells):
                            self.create_rectangle(fc + 1, 0, sc, self.current_height - 1, fill = self.selected_cells_background, outline = "", tag = "s")
                            tf = self.selected_cells_foreground
                        elif c in self.highlighted_cells:
                            self.create_rectangle(fc + 1, 0, sc, self.current_height - 1, fill = self.highlighted_cells[c][0], outline = "", tag = "s")
                            tf = self.text_color if self.highlighted_cells[c][1] is None else self.highlighted_cells[c][1]
                        else:
                            tf = self.text_color
                        mw = sc - fc - 5
                        x = fc + 7
                        if x > x_stop:
                            continue
                        if isinstance(self.MT.my_hdrs, int):
                            try:
                                lns = self.MT.data_ref[self.MT.my_hdrs][c].split("\n")
                            except:
                                lns = (f"{c + 1}", )
                        else:
                            try:
                                lns = self.MT.my_hdrs[c].split("\n")
                            except:
                                lns = (f"{c + 1}", )
                        y = self.MT.hdr_fl_ins
                        if incfl:
                            fl = lns[0]
                            t = self.create_text(x, y, text = fl, fill = tf, font = self.MT.my_hdr_font, anchor = "w", tag = "t")
                            wd = self.bbox(t)
                            wd = wd[2] - wd[0]
                            if wd > mw:
                                nl = int(len(fl) * (mw / wd)) - 1
                                self.itemconfig(t, text = fl[:nl])
                                wd = self.bbox(t)
                                while wd[2] - wd[0] > mw:
                                    nl -= 1
                                    self.dchars(t, nl)
                                    wd = self.bbox(t)
                        if len(lns) > 1:
                            stl = int((top - y) / self.MT.hdr_xtra_lines_increment) - 1
                            if stl < 1:
                                stl = 1
                            y += (stl * self.MT.hdr_xtra_lines_increment)
                            if y + self.MT.hdr_half_txt_h < self.current_height:
                                for i in range(stl, len(lns)):
                                    txt = lns[i]
                                    t = self.create_text(x, y, text = txt, fill = tf, font = self.MT.my_hdr_font, anchor = "w", tag = "t")
                                    wd = self.bbox(t)
                                    wd = wd[2] - wd[0]
                                    if wd > mw:
                                        nl = int(len(txt) * (mw / wd)) - 1
                                        self.itemconfig(t, text = txt[:nl])
                                        wd = self.bbox(t)
                                        while wd[2] - wd[0] > mw:
                                            nl -= 1
                                            self.dchars(t, nl)
                                            wd = self.bbox(t)
                                    y += self.MT.hdr_xtra_lines_increment
                                    if y + self.MT.hdr_half_txt_h > self.current_height:
                                        break
            else:
                if self.align == "center":
                    for c in range(start_col, end_col - 1):
                        fc = self.MT.col_positions[c]
                        sc = self.MT.col_positions[c + 1]
                        if self.MT.displayed_columns[c] in self.highlighted_cells and (c in self.selected_cells or c in self.MT.selected_cols):
                            c_1 = self.highlighted_cells[self.MT.displayed_columns[c]][0] if self.highlighted_cells[self.MT.displayed_columns[c]][0].startswith("#") else Color_Map_[self.highlighted_cells[self.MT.displayed_columns[c]][0]]
                            self.create_rectangle(fc + 1,
                                                  0,
                                                  sc,
                                                  self.current_height - 1,
                                                  fill = (f"#{int((int(c_1[1:3], 16) + int(c_2[1:3], 16)) / 2):02X}" +
                                                          f"{int((int(c_1[3:5], 16) + int(c_2[3:5], 16)) / 2):02X}" +
                                                          f"{int((int(c_1[5:], 16) + int(c_2[5:], 16)) / 2):02X}"),
                                                  outline = "",
                                                  tag = "s")
                            tf = self.selected_cells_foreground if self.highlighted_cells[c][1] is None else self.highlighted_cells[c][1]
                        elif c in (self.MT.selected_cols or self.selected_cells):
                            self.create_rectangle(fc + 1, 0, sc, self.current_height - 1, fill = self.selected_cells_background, outline = "", tag = "s")
                            tf = self.selected_cells_foreground
                        elif self.MT.displayed_columns[c] in self.highlighted_cells:
                            self.create_rectangle(fc + 1, 0, sc, self.current_height - 1, fill = self.highlighted_cells[self.MT.displayed_columns[c]][0], outline = "", tag = "s")
                            tf = self.text_color if self.highlighted_cells[self.MT.displayed_columns[c]][1] is None else self.highlighted_cells[self.MT.displayed_columns[c]][1]
                        else:
                            tf = self.text_color
                        if fc + 7 > x_stop:
                            continue
                        mw = sc - fc - 5
                        x = fc + floor(mw / 2)
                        if isinstance(self.MT.my_hdrs, int):
                            try:
                                lns = self.MT.data_ref[self.MT.my_hdrs][c].split("\n")
                            except:
                                lns = (f"{c + 1}", )
                        else:
                            try:
                                lns = self.MT.my_hdrs[self.MT.displayed_columns[c]].split("\n")
                            except:
                                lns = (f"{c + 1}", )
                        y = self.MT.hdr_fl_ins
                        if incfl:
                            fl = lns[0]
                            t = self.create_text(x, y, text = fl, fill = tf, font = self.MT.my_hdr_font, anchor = "center", tag = "t")
                            wd = self.bbox(t)
                            wd = wd[2] - wd[0]
                            if wd > mw:
                                tl = len(fl)
                                slce = tl - floor(tl * (mw / wd))
                                if slce % 2:
                                    slce += 1
                                else:
                                    slce += 2
                                slce = int(slce / 2)
                                fl = fl[slce:tl - slce]
                                self.itemconfig(t, text = fl)
                                wd = self.bbox(t)
                                while wd[2] - wd[0] > mw:
                                    fl = fl[1: - 1]
                                    self.itemconfig(t, text = fl)
                                    wd = self.bbox(t)
                        if len(lns) > 1:
                            stl = int((top - y) / self.MT.hdr_xtra_lines_increment) - 1
                            if stl < 1:
                                stl = 1
                            y += (stl * self.MT.hdr_xtra_lines_increment)
                            if y + self.MT.hdr_half_txt_h < self.current_height:
                                for i in range(stl, len(lns)):
                                    txt = lns[i]
                                    t = self.create_text(x, y, text = txt, fill = tf, font = self.MT.my_hdr_font, anchor = "center", tag = "t")
                                    wd = self.bbox(t)
                                    wd = wd[2] - wd[0]
                                    if wd > mw:
                                        tl = len(txt)
                                        slce = tl - floor(tl * (mw / wd))
                                        if slce % 2:
                                            slce += 1
                                        else:
                                            slce += 2
                                        slce = int(slce / 2)
                                        txt = txt[slce:tl - slce]
                                        self.itemconfig(t, text = txt)
                                        wd = self.bbox(t)
                                        while wd[2] - wd[0] > mw:
                                            txt = txt[1: - 1]
                                            self.itemconfig(t, text = txt)
                                            wd = self.bbox(t)
                                    y += self.MT.hdr_xtra_lines_increment
                                    if y + self.MT.hdr_half_txt_h > self.current_height:
                                        break
                elif self.align == "w":
                    for c in range(start_col, end_col - 1):
                        fc = self.MT.col_positions[c]
                        sc = self.MT.col_positions[c + 1]
                        if self.MT.displayed_columns[c] in self.highlighted_cells and (c in self.selected_cells or c in self.MT.selected_cols):
                            c_1 = self.highlighted_cells[self.MT.displayed_columns[c]][0] if self.highlighted_cells[self.MT.displayed_columns[c]][0].startswith("#") else Color_Map_[self.highlighted_cells[self.MT.displayed_columns[c]][0]]
                            self.create_rectangle(fc + 1,
                                                  0,
                                                  sc,
                                                  self.current_height - 1,
                                                  fill = (f"#{int((int(c_1[1:3], 16) + int(c_2[1:3], 16)) / 2):02X}" +
                                                          f"{int((int(c_1[3:5], 16) + int(c_2[3:5], 16)) / 2):02X}" +
                                                          f"{int((int(c_1[5:], 16) + int(c_2[5:], 16)) / 2):02X}"),
                                                  outline = "",
                                                  tag = "s")
                            tf = self.selected_cells_foreground if self.highlighted_cells[c][1] is None else self.highlighted_cells[c][1]
                        elif c in (self.MT.selected_cols or self.selected_cells):
                            self.create_rectangle(fc + 1, 0, sc, self.current_height - 1, fill = self.selected_cells_background, outline = "", tag = "s")
                            tf = self.selected_cells_foreground
                        elif self.MT.displayed_columns[c] in self.highlighted_cells:
                            self.create_rectangle(fc + 1, 0, sc, self.current_height - 1, fill = self.highlighted_cells[self.MT.displayed_columns[c]][0], outline = "", tag = "s")
                            tf = self.text_color if self.highlighted_cells[self.MT.displayed_columns[c]][1] is None else self.highlighted_cells[self.MT.displayed_columns[c]][1]
                        else:
                            tf = self.text_color
                        mw = sc - fc - 5
                        x = fc + 7
                        if x > x_stop:
                            continue
                        if isinstance(self.MT.my_hdrs, int):
                            try:
                                lns = self.MT.data_ref[self.MT.my_hdrs][c].split("\n")
                            except:
                                lns = (f"{c + 1}", )
                        else:
                            try:
                                lns = self.MT.my_hdrs[self.MT.displayed_columns[c]].split("\n")
                            except:
                                lns = (f"{c + 1}", )
                        y = self.MT.hdr_fl_ins
                        if incfl:
                            fl = lns[0]
                            t = self.create_text(x, y, text = fl, fill = tf, font = self.MT.my_hdr_font, anchor = "w", tag = "t")
                            wd = self.bbox(t)
                            wd = wd[2] - wd[0]
                            if wd > mw:
                                nl = int(len(fl) * (mw / wd)) - 1
                                self.itemconfig(t, text = fl[:nl])
                                wd = self.bbox(t)
                                while wd[2] - wd[0] > mw:
                                    nl -= 1
                                    self.dchars(t, nl)
                                    wd = self.bbox(t)
                        if len(lns) > 1:
                            stl = int((top - y) / self.MT.hdr_xtra_lines_increment) - 1
                            if stl < 1:
                                stl = 1
                            y += (stl * self.MT.hdr_xtra_lines_increment)
                            if y + self.MT.hdr_half_txt_h < self.current_height:
                                for i in range(stl, len(lns)):
                                    txt = lns[i]
                                    t = self.create_text(x, y, text = txt, fill = tf, font = self.MT.my_hdr_font, anchor = "w", tag = "t")
                                    wd = self.bbox(t)
                                    wd = wd[2] - wd[0]
                                    if wd > mw:
                                        nl = int(len(txt) * (mw / wd)) - 1
                                        self.itemconfig(t, text = txt[:nl])
                                        wd = self.bbox(t)
                                        while wd[2] - wd[0] > mw:
                                            nl -= 1
                                            self.dchars(t, nl)
                                            wd = self.bbox(t)
                                    y += self.MT.hdr_xtra_lines_increment
                                    if y + self.MT.hdr_half_txt_h > self.current_height:
                                        break
            self.create_line(x1, self.current_height - 1, x_stop, self.current_height - 1, fill = self.header_border_color, width = 1, tag = "h")
        except:
            return
        
    def GetCellCoords(self, event = None, r = None, c = None):
        pass


class RowIndexes(tk.Canvas):
    def __init__(self,
                 parentframe = None,
                 main_canvas = None,
                 header_canvas = None,
                 max_rh = None,
                 max_row_width = None,
                 row_index_align = None,
                 row_index_width = None,
                 row_index_background = None,
                 row_index_border_color = None,
                 row_index_grid_color = None,
                 row_index_foreground = None,
                 row_index_select_background = None,
                 row_index_select_foreground = None,
                 drag_and_drop_color = None,
                 resizing_line_color = None):
        tk.Canvas.__init__(self,
                           parentframe,
                           height = None,
                           background = row_index_background,
                           highlightthickness = 0)
        self.parentframe = parentframe
        self.extra_motion_func = None
        self.extra_b1_press_func = None
        self.extra_b1_motion_func = None
        self.extra_b1_release_func = None
        self.selection_binding_func = None
        self.drag_selection_binding_func = None
        self.ri_extra_drag_drop_func = None
        self.extra_double_b1_func = None
        self.new_row_width = 0
        if row_index_width is None:
            self.set_width(100)
            self.default_width = 100
        else:
            self.set_width(row_index_width)
            self.default_width = row_index_width
        self.max_rh = float(max_rh)
        self.max_row_width = float(max_row_width)
        self.MT = main_canvas         # is set from within MainTable() __init__
        self.CH = header_canvas      # is set from within MainTable() __init__
        self.TL = None                # is set from within TopLeftRectangle() __init__
        self.text_color = row_index_foreground
        self.grid_color = row_index_grid_color
        self.row_index_border_color = row_index_border_color
        self.selected_cells_background = row_index_select_background
        self.selected_cells_foreground = row_index_select_foreground
        self.row_index_background = row_index_background
        self.drag_and_drop_color = drag_and_drop_color
        self.resizing_line_color = resizing_line_color
        self.align = row_index_align
        self.selected_cells = defaultdict(int)
        self.highlighted_cells = {}
        self.drag_and_drop_enabled = False
        self.dragged_row = None
        self.width_resizing_enabled = False
        self.height_resizing_enabled = False
        self.double_click_resizing_enabled = False
        self.row_selection_enabled = False
        self.visible_row_dividers = []
        self.row_width_resize_bbox = tuple()
        self.rsz_w = None
        self.rsz_h = None
        self.currently_resizing_width = False
        self.currently_resizing_height = False
        self.bind("<Motion>", self.mouse_motion)
        self.bind("<ButtonPress-1>", self.b1_press)
        self.bind("<Shift-ButtonPress-1>",self.shift_b1_press)
        self.bind("<B1-Motion>", self.b1_motion)
        self.bind("<ButtonRelease-1>", self.b1_release)
        self.bind("<Double-Button-1>", self.double_b1)
        self.bind("<MouseWheel>", self.mousewheel)

    def basic_bindings(self, onoff = "enable"):
        if onoff == "enable":
            self.bind("<Motion>", self.mouse_motion)
            self.bind("<ButtonPress-1>", self.b1_press)
            self.bind("<B1-Motion>", self.b1_motion)
            self.bind("<ButtonRelease-1>", self.b1_release)
            self.bind("<Double-Button-1>", self.double_b1)
            self.bind("<MouseWheel>", self.mousewheel)
        elif onoff == "disable":
            self.unbind("<Motion>")
            self.unbind("<ButtonPress-1>")
            self.unbind("<B1-Motion>")
            self.unbind("<ButtonRelease-1>")
            self.unbind("<Double-Button-1>")
            self.unbind("<MouseWheel>")

    def mousewheel(self, event = None):
        if event.num == 5 or event.delta == -120:
            self.yview_scroll(1, "units")
            self.MT.yview_scroll(1, "units")
        if event.num == 4 or event.delta == 120:
            if self.canvasy(0) <= 0:
                return
            self.yview_scroll( - 1, "units")
            self.MT.yview_scroll( - 1, "units")
        self.MT.main_table_redraw_grid_and_text(redraw_row_index = True)

    def set_width(self, new_width, set_TL = False):
        self.current_width = new_width
        self.config(width = new_width)
        if set_TL:
            self.TL.set_dimensions(new_w = new_width)

    def enable_bindings(self, binding):
        if binding == "row_width_resize":
            self.width_resizing_enabled = True
        elif binding == "row_height_resize":
            self.height_resizing_enabled = True
        elif binding == "double_click_row_resize":
            self.double_click_resizing_enabled = True
        elif binding == "row_select":
            self.row_selection_enabled = True
        elif binding == "drag_and_drop":
            self.drag_and_drop_enabled = True

    def disable_bindings(self, binding):
        if binding == "row_width_resize":
            self.width_resizing_enabled = False
        elif binding == "row_height_resize":
            self.height_resizing_enabled = False
        elif binding == "double_click_row_resize":
            self.double_click_resizing_enabled = False
        elif binding == "row_select":
            self.row_selection_enabled = False
        elif binding == "drag_and_drop":
            self.drag_and_drop_enabled = False

    def check_mouse_position_height_resizers(self, x, y):
        ov = None
        for x1, y1, x2, y2 in self.visible_row_dividers:
            if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                ov = self.find_overlapping(x1, y1, x2, y2)
                break
        return ov

    def shift_b1_press(self, event):
        y = event.y
        r = self.MT.identify_row(y = y)
        if self.drag_and_drop_enabled or self.row_selection_enabled and self.rsz_h is None and self.rsz_w is None:
            if r < len(self.MT.row_positions) - 1:
                if r not in self.MT.selected_rows and self.row_selection_enabled:
                    r = int(r)
                    if self.MT.currently_selected and self.MT.currently_selected[0] == "row":
                        min_r = int(self.MT.currently_selected[1])
                        self.selected_cells = defaultdict(int)
                        self.CH.selected_cells = defaultdict(int)
                        self.MT.selected_cols = set()
                        self.MT.selected_rows = set()
                        self.MT.selected_cells = set()
                        if r > min_r:
                            for i in range(min_r, r + 1):
                                self.selected_cells[i] += 1
                                self.MT.selected_rows.add(i)
                        elif r < min_r:
                            for i in range(r, min_r + 1):
                                self.selected_cells[i] += 1
                                self.MT.selected_rows.add(i)
                    else:
                        self.select_row(r)
                    self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
                    if self.selection_binding_func is not None:
                        self.selection_binding_func(("row", r))
                elif r in self.MT.selected_rows:
                    self.dragged_row = r

    def mouse_motion(self, event):
        if not self.currently_resizing_height and not self.currently_resizing_width:
            x = self.canvasx(event.x)
            y = self.canvasy(event.y)
            mouse_over_resize = False
            if self.height_resizing_enabled and not mouse_over_resize:
                ov = self.check_mouse_position_height_resizers(x, y)
                if ov is not None:
                    for itm in ov:
                        tgs = self.gettags(itm)
                        if "h" == tgs[0]:
                            break
                    r = int(tgs[1])
                    self.config(cursor = "sb_v_double_arrow")
                    self.rsz_h = r
                    mouse_over_resize = True
                else:
                    self.rsz_h = None
            if self.width_resizing_enabled and not mouse_over_resize:
                try:
                    x1, y1, x2, y2 = self.row_width_resize_bbox[0], self.row_width_resize_bbox[1], self.row_width_resize_bbox[2], self.row_width_resize_bbox[3]
                    if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                        self.config(cursor = "sb_h_double_arrow")
                        self.rsz_w = True
                        mouse_over_resize = True
                    else:
                        self.rsz_w = None
                except:
                    self.rsz_w = None
            if not mouse_over_resize:
                self.MT.reset_mouse_motion_creations()
        if self.extra_motion_func is not None:
            self.extra_motion_func(event)
        
    def b1_press(self, event = None):
        self.focus_set()
        self.MT.unbind("<MouseWheel>")
        x = self.canvasx(event.x)
        y = self.canvasy(event.y)
        if self.check_mouse_position_height_resizers(x, y) is None:
            self.rsz_h = None
        if not x >= self.row_width_resize_bbox[0] and y >= self.row_width_resize_bbox[1] and x <= self.row_width_resize_bbox[2] and y <= self.row_width_resize_bbox[3]:
            self.rsz_w = None
        if self.height_resizing_enabled and self.rsz_h is not None:
            self.currently_resizing_height = True
            y = self.MT.row_positions[self.rsz_h]
            line2y = self.MT.row_positions[self.rsz_h - 1]
            x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
            self.create_line(0, y, self.current_width, y, width = 1, fill = self.resizing_line_color, tag = "rhl")
            self.MT.create_line(x1, y, x2, y, width = 1, fill = self.resizing_line_color, tag = "rhl")
            self.create_line(0, line2y, self.current_width, line2y, width = 1, fill = self.resizing_line_color, tag = "rhl2")
            self.MT.create_line(x1, line2y, x2, line2y, width = 1, fill = self.resizing_line_color, tag = "rhl2")
        elif self.width_resizing_enabled and self.rsz_h is None and self.rsz_w == True:
            self.currently_resizing_width = True
            x1, y1, x2, y2 = self.MT.get_canvas_visible_area()
            x = int(event.x)
            if x < self.MT.min_cw:
                x = int(self.MT.min_cw)
            self.new_row_width = x
            self.create_line(x, y1, x, y2, width = 1, fill = self.resizing_line_color, tag = "rwl")
        elif self.MT.identify_row(y = event.y, allow_end = False) is None:
            self.MT.deselect("all")
        elif self.row_selection_enabled and self.rsz_h is None and self.rsz_w is None:
            r = self.MT.identify_row(y = event.y)
            if r < len(self.MT.row_positions) - 1:
                self.select_row(r, redraw = True)
        if self.extra_b1_press_func is not None:
            self.extra_b1_press_func(event)
    
    def b1_motion(self, event):
        x1,y1,x2,y2 = self.MT.get_canvas_visible_area()
        if self.height_resizing_enabled and self.rsz_h is not None and self.currently_resizing_height:
            y = self.canvasy(event.y)
            size = y - self.MT.row_positions[self.rsz_h - 1]
            if not size <= self.MT.min_rh and size < self.max_rh:
                self.delete("rhl")
                self.MT.delete("rhl")
                self.create_line(0, y, self.current_width, y, width = 1, fill = self.resizing_line_color, tag = "rhl")
                self.MT.create_line(x1, y, x2, y, width = 1, fill = self.resizing_line_color, tag = "rhl")
        elif self.width_resizing_enabled and self.rsz_w is not None and self.currently_resizing_width:
            evx = event.x
            self.delete("rwl")
            self.MT.delete("rwl")
            if evx > self.current_width:
                x = self.MT.canvasx(evx - self.current_width)
                if evx > self.max_row_width:
                    evx = int(self.max_row_width)
                    x = self.MT.canvasx(evx - self.current_width)
                self.new_row_width = evx
                self.MT.create_line(x, y1, x, y2, width = 1, fill = self.resizing_line_color, tag = "rwl")
            else:
                x = evx
                if x < self.MT.min_cw:
                    x = int(self.MT.min_cw)
                self.new_row_width = x
                self.create_line(x, y1, x, y2, width = 1, fill = self.resizing_line_color, tag = "rwl")
        if self.drag_and_drop_enabled and self.row_selection_enabled and self.rsz_h is None and self.rsz_w is None and self.dragged_row is not None and self.MT.selected_rows:
            y = self.canvasy(event.y)
            if y > 0 and y < self.MT.row_positions[-1]:
                y = event.y
                hend = self.winfo_height()
                if y >= hend - 0:
                    end_row = bisect.bisect_right(self.MT.row_positions, self.canvasy(hend))
                    end_row -= 1
                    if not end_row == len(self.MT.row_positions) - 1:
                        try:
                            self.MT.see(r = end_row, c = 0, keep_yscroll = False, keep_xscroll = True, bottom_right_corner = False, check_cell_visibility = True)
                        except:
                            pass
                elif y <= 0:
                    start_row = bisect.bisect_left(self.MT.row_positions, self.canvasy(0))
                    if y <= -40:
                        start_row -= 3
                    else:
                        start_row -= 2
                    if start_row <= 0:
                        start_row = 0
                    try:
                        self.MT.see(r = start_row, c = 0, keep_yscroll = False, keep_xscroll = True, bottom_right_corner = False, check_cell_visibility = True)
                    except:
                        pass
                rectw = self.MT.row_positions[max(self.MT.selected_rows) + 1] - self.MT.row_positions[min(self.MT.selected_rows)]
                start = self.canvasy(event.y - int(rectw / 2))
                end = self.canvasy(event.y + int(rectw / 2))
                self.delete("dd")
                self.create_rectangle(0, start, self.current_width - 1, end, fill = self.drag_and_drop_color, outline = self.grid_color, tag = "dd")
                self.tag_raise("dd")
                self.tag_raise("t")
                self.tag_raise("h")
        elif self.MT.drag_selection_enabled and self.row_selection_enabled and self.rsz_h is None and self.rsz_w is None:
            end_row = self.MT.identify_row(y = event.y)
            if end_row < len(self.MT.row_positions) - 1 and len(self.MT.currently_selected) == 2:
                if self.MT.currently_selected[0] == "row":
                    start_row = self.MT.currently_selected[1]
                    self.MT.selected_cols = set()
                    self.MT.selected_rows = set()
                    self.MT.selected_cells = set()
                    self.CH.selected_cells = defaultdict(int)
                    self.selected_cells = defaultdict(int)
                    if end_row >= start_row:
                        for r in range(start_row, end_row + 1):
                            self.selected_cells[r] += 1
                            self.MT.selected_rows.add(r)
                    elif end_row < start_row:
                        for r in range(end_row, start_row + 1):
                            self.selected_cells[r] += 1
                            self.MT.selected_rows.add(r)
                    if self.drag_selection_binding_func is not None:
                        self.drag_selection_binding_func(("rows", sorted([start_row, end_row])))
            if event.y > self.winfo_height():
                try:
                    self.MT.yview_scroll(1, "units")
                    self.yview_scroll(1, "units")
                except:
                    pass
            elif event.y < 0 and self.canvasy(self.winfo_height()) > 0:
                try:
                    self.yview_scroll(-1, "units")
                    self.MT.yview_scroll(-1, "units")
                except:
                    pass
            self.MT.main_table_redraw_grid_and_text(redraw_header = False, redraw_row_index = True)
        if self.extra_b1_motion_func is not None:
            self.extra_b1_motion_func(event)
            
    def b1_release(self, event = None):
        self.MT.bind("<MouseWheel>", self.MT.mousewheel)
        if self.height_resizing_enabled and self.rsz_h is not None and self.currently_resizing_height:
            self.currently_resizing_height = False
            new_row_pos = self.coords("rhl")[1]
            self.delete("rhl", "rhl2")
            self.MT.delete("rhl", "rhl2")
            increment = new_row_pos - self.MT.row_positions[self.rsz_h]
            self.MT.row_positions[self.rsz_h + 1:] = [e + increment for e in islice(self.MT.row_positions, self.rsz_h + 1, len(self.MT.row_positions))]
            self.MT.row_positions[self.rsz_h] = new_row_pos
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        elif self.width_resizing_enabled and self.rsz_w is not None and self.currently_resizing_width:
            self.currently_resizing_width = False
            self.delete("rwl")
            self.MT.delete("rwl")
            self.set_width(self.new_row_width, set_TL = True)
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.drag_and_drop_enabled and self.MT.selected_rows and self.row_selection_enabled and self.rsz_h is None and self.rsz_w is None and self.dragged_row is not None:
            self.delete("dd")
            y = event.y
            r = self.MT.identify_row(y = y)
            if r != self.dragged_row and r is not None and len(self.MT.selected_rows) != (len(self.MT.row_positions) - 1):
                rowsiter = list(self.MT.selected_rows)
                rowsiter.sort()
                stins = rowsiter[0]
                endins = rowsiter[-1] + 1
                if self.dragged_row < r and r >= len(self.MT.row_positions) - 1:
                    r -= 1
                r_ = int(r)
                if r >= endins:
                    r += 1
                if self.ri_extra_drag_drop_func is not None:
                    self.ri_extra_drag_drop_func(self.MT.selected_rows, int(r_))
                else:
                    if stins > r:
                        self.MT.data_ref[r:r] = self.MT.data_ref[stins:endins]
                        self.MT.data_ref[stins + len(rowsiter):endins + len(rowsiter)] = []
                        if not isinstance(self.MT.my_row_index, int) and self.MT.my_row_index:
                            try:
                                self.MT.my_row_index[r:r] = self.MT.my_row_index[stins:endins]
                                self.MT.my_row_index[stins + len(rowsiter):endins + len(rowsiter)] = []
                            except:
                                pass
                    else:
                        self.MT.data_ref[r:r] = self.MT.data_ref[stins:endins]
                        self.MT.data_ref[stins:endins] = []
                        if not isinstance(self.MT.my_row_index, int) and self.MT.my_row_index:
                            try:
                                self.MT.my_row_index[r:r] = self.MT.my_row_index[stins:endins]
                                self.MT.my_row_index[stins:endins] = []
                            except:
                                pass
                rhs = self.MT.parentframe.get_row_heights()
                if stins > r:
                    rhs[r:r] = rhs[stins:endins]
                    rhs[stins + len(rowsiter):endins + len(rowsiter)] = []
                else:
                    rhs[r:r] = rhs[stins:endins]
                    rhs[stins:endins] = []
                self.MT.parentframe.set_row_heights(rhs)
                if (r_ - 1) + len(rowsiter) > len(self.MT.row_positions) - 1:
                    sels_start = len(self.MT.row_positions) - 1 - len(rowsiter)
                    newrowidxs = tuple(range(sels_start, len(self.MT.row_positions) - 1))
                else:
                    if r_ > endins:
                        r_ += 1
                        sels_start = r_ - len(rowsiter)
                    else:
                        if r_ == endins and len(rowsiter) == 1:
                            pass
                        else:
                            if r_ > endins:
                                r_ += 1
                            if r_ == endins:
                                r_ -= 1
                            if r_ < 0:
                                r_ = 0
                        sels_start = r_
                    newrowidxs = tuple(range(sels_start, sels_start + len(rowsiter)))
                self.MT.selected_rows = set()
                self.MT.selected_cells = set()
                self.selected_cells = defaultdict(int)
                self.CH.selected_cells = defaultdict(int)
                self.MT.selected_cols = set()
                for rowsel in newrowidxs:
                    self.MT.selected_rows.add(rowsel)
                    self.selected_cells[rowsel] += 1
                self.MT.undo_storage = deque(maxlen = 20)
                self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        self.dragged_row = None
        self.currently_resizing_width = False
        self.currently_resizing_height = False
        self.rsz_w = None
        self.rsz_h = None
        self.mouse_motion(event)
        if self.extra_b1_release_func is not None:
            self.extra_b1_release_func(event)

    def double_b1(self, event = None):
        self.focus_set()
        if self.double_click_resizing_enabled and self.height_resizing_enabled and self.rsz_h is not None and not self.currently_resizing_height:
            row = self.rsz_h - 1
            self.set_row_height(row)
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        elif self.row_selection_enabled and self.rsz_h is None and self.rsz_w is None:
            r = self.MT.identify_row(y = event.y)
            if r < len(self.MT.row_positions) - 1:
                self.select_row(r, redraw = True)
        self.mouse_motion(event)
        self.rsz_h = None
        if self.extra_double_b1_func is not None:
            self.extra_double_b1_func(event)

    def set_row_height(self, row, only_set_if_too_small = False):
        r_norm = row + 1
        r_extra = row + 2
        try:
            new_height = self.GetLinesHeight(max((cll for cll in self.MT.data_ref[row]), key = self.GetNumLines))
        except:
            new_height = int(self.MT.min_rh)
        if new_height < self.MT.min_rh:
            new_height = int(self.MT.min_rh)
        elif new_height > self.max_rh:
            new_height = int(self.max_rh)
        if only_set_if_too_small:
            if new_height <= self.MT.row_positions[row + 1] - self.MT.row_positions[row]:
                return
        new_row_pos = self.MT.row_positions[row] + new_height
        increment = new_row_pos - self.MT.row_positions[r_norm]
        self.MT.row_positions[r_extra:] = [e + increment for e in islice(self.MT.row_positions, r_extra, len(self.MT.row_positions))]
        self.MT.row_positions[r_norm] = new_row_pos

    def GetNumLines(self, cll):
        return len(cll.split("\n"))

    def GetLinesHeight(self, cll):
        lns = cll.split("\n")
        if len(lns) > 1:
            y = self.MT.fl_ins
            for i in range(len(lns)):
                y += self.MT.xtra_lines_increment
        else:
            y = int(self.MT.min_rh)
        return y

    def highlight_cells(self, r = 0, cells = tuple(), bg = None, fg = None, redraw = False):
        if bg is None and fg is None:
            return
        if cells:
            self.highlighted_cells = {r_: (bg, fg)  for r_ in cells}
        else:
            self.highlighted_cells[r] = (bg, fg)
        if redraw:
            self.MT.main_table_redraw_grid_and_text(False, True)

    def add_selection(self, r, redraw = False, run_binding_func = True):
        r = int(r)
        self.MT.currently_selected = ("row", r)
        self.selected_cells[r] += 1
        self.MT.selected_rows.add(r)
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.selection_binding_func is not None and run_binding_func:
            self.selection_binding_func(("row", r))

    def select_row(self, r, redraw = False):
        r = int(r)
        self.selected_cells = defaultdict(int)
        self.selected_cells[r] += 1
        self.CH.selected_cells = defaultdict(int)
        self.MT.selected_cols = set()
        self.MT.selected_rows = {r}
        self.MT.selected_cells = set()
        self.MT.currently_selected = ("row", r)
        if redraw:
            self.MT.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.selection_binding_func is not None:
            self.selection_binding_func(("row", r))

    def redraw_grid_and_text(self, last_row_line_pos, y1, y_stop, start_row, end_row, y2, x1, x_stop):
        try:
            self.configure(scrollregion = (0, 0, self.current_width, last_row_line_pos + 100))
            self.delete("fh", "h", "v", "t", "s")
            self.visible_row_dividers = []
            y = self.MT.row_positions[start_row]
            self.create_line(0, y, self.current_width, y, fill = self.grid_color, width = 1, tag = "fh")
            xend = self.current_width - 6
            self.row_width_resize_bbox = (self.current_width - 5, y1, self.current_width, y2)
            if self.height_resizing_enabled:
                for r in range(start_row + 1,end_row):
                    y = self.MT.row_positions[r]
                    self.visible_row_dividers.append((1, y - 4, xend, y + 4))
                    self.create_line(0, y, self.current_width, y, fill = self.grid_color, width = 1, tag = ("h", f"{r}"))
            else:
                for r in range(start_row + 1,end_row):
                    y = self.MT.row_positions[r]
                    self.create_line(0, y, self.current_width, y, fill = self.grid_color, width = 1, tag = ("h", f"{r}"))
            sb = y2 + 2
            c_2 = self.selected_cells_background if self.selected_cells_background.startswith("#") else Color_Map_[self.selected_cells_background]
            if self.align == "center":
                mw = self.current_width - 7
                x = floor(mw / 2)
                for r in range(start_row, end_row - 1):
                    fr = self.MT.row_positions[r]
                    sr = self.MT.row_positions[r+1]
                    if sr > sb:
                        sr = sb
                    if r in self.highlighted_cells and (r in self.selected_cells or r in self.MT.selected_rows):
                        c_1 = self.highlighted_cells[r][0] if self.highlighted_cells[r][0].startswith("#") else Color_Map_[self.highlighted_cells[r][0]]
                        self.create_rectangle(0,
                                              fr + 1,
                                              self.current_width - 1,
                                              sr,
                                              fill = (f"#{int((int(c_1[1:3], 16) + int(c_2[1:3], 16)) / 2):02X}" +
                                                      f"{int((int(c_1[3:5], 16) + int(c_2[3:5], 16)) / 2):02X}" +
                                                      f"{int((int(c_1[5:], 16) + int(c_2[5:], 16)) / 2):02X}"),
                                              outline = "",
                                              tag = "s")
                        tf = self.selected_cells_foreground if self.highlighted_cells[r][1] is None else self.highlighted_cells[r][1]
                    elif r in (self.selected_cells or self.MT.selected_rows):
                        self.create_rectangle(0, fr + 1, self.current_width - 1, sr, fill = self.selected_cells_background, outline = "", tag = "s")
                        tf = self.selected_cells_foreground
                    elif r in self.highlighted_cells:
                        self.create_rectangle(0, fr + 1, self.current_width - 1, sr, fill = self.highlighted_cells[r][0], outline = "", tag = "s")
                        tf = self.text_color if self.highlighted_cells[r][1] is None else self.highlighted_cells[r][1]
                    else:
                        tf = self.text_color
                    if isinstance(self.MT.my_row_index, int):
                        lns = self.MT.data_ref[r][self.MT.my_row_index].split("\n")
                    else:
                        try:
                            lns = self.MT.my_row_index[r].split("\n")
                        except:
                            lns = (f"{r + 1}", )
                    fl = lns[0]
                    y = fr + self.MT.fl_ins
                    if y + self.MT.half_txt_h > y1:
                        t = self.create_text(x, y, text = fl, fill = tf, font = self.MT.my_font, anchor = "center", tag = "t")
                        wd = self.bbox(t)
                        wd = wd[2] - wd[0]
                        if wd > mw:
                            tl = len(fl)
                            slce = tl - floor(tl * (mw / wd))
                            if slce % 2:
                                slce += 1
                            else:
                                slce += 2
                            slce = int(slce / 2)
                            fl = fl[slce:tl - slce]
                            self.itemconfig(t, text = fl)
                            wd = self.bbox(t)
                            while wd[2] - wd[0] > mw:
                                fl = fl[1: - 1]
                                self.itemconfig(t, text = fl)
                                wd = self.bbox(t)
                    if len(lns) > 1:
                        stl = int((y1 - y) / self.MT.xtra_lines_increment) - 1
                        if stl < 1:
                            stl = 1
                        y += (stl * self.MT.xtra_lines_increment)
                        if y + self.MT.half_txt_h < sr:
                            for i in range(stl,len(lns)):
                                txt = lns[i]
                                t = self.create_text(x, y, text = txt, fill = tf, font = self.MT.my_font, anchor = "center", tag = "t")
                                wd = self.bbox(t)
                                wd = wd[2] - wd[0]
                                if wd > mw:
                                    tl = len(txt)
                                    slce = tl - floor(tl * (mw / wd))
                                    if slce % 2:
                                        slce += 1
                                    else:
                                        slce += 2
                                    slce = int(slce / 2)
                                    txt = txt[slce:tl - slce]
                                    self.itemconfig(t, text = txt)
                                    wd = self.bbox(t)
                                    while wd[2] - wd[0] > mw:
                                        txt = txt[1: - 1]
                                        self.itemconfig(t, text = txt)
                                        wd = self.bbox(t)
                                y += self.MT.xtra_lines_increment
                                if y + self.MT.half_txt_h > sr:
                                    break
            elif self.align == "w":
                mw = self.current_width - 7
                x = 7
                for r in range(start_row,end_row - 1):
                    fr = self.MT.row_positions[r]
                    sr = self.MT.row_positions[r + 1]
                    if sr > sb:
                        sr = sb
                    if r in self.highlighted_cells and (r in self.selected_cells or r in self.MT.selected_rows):
                        c_1 = self.highlighted_cells[r][0] if self.highlighted_cells[r][0].startswith("#") else Color_Map_[self.highlighted_cells[r][0]]
                        self.create_rectangle(0,
                                              fr + 1,
                                              self.current_width - 1,
                                              sr,
                                              fill = (f"#{int((int(c_1[1:3], 16) + int(c_2[1:3], 16)) / 2):02X}" +
                                                      f"{int((int(c_1[3:5], 16) + int(c_2[3:5], 16)) / 2):02X}" +
                                                      f"{int((int(c_1[5:], 16) + int(c_2[5:], 16)) / 2):02X}"),
                                              outline = "",
                                              tag = "s")
                        tf = self.selected_cells_foreground if self.highlighted_cells[r][1] is None else self.highlighted_cells[r][1]
                    elif r in (self.selected_cells or self.MT.selected_rows):
                        self.create_rectangle(0, fr + 1, self.current_width - 1, sr, fill = self.selected_cells_background, outline = "", tag = "s")
                        tf = self.selected_cells_foreground
                    elif r in self.highlighted_cells:
                        self.create_rectangle(0, fr + 1, self.current_width - 1, sr, fill = self.highlighted_cells[r][0], outline = "", tag = "s")
                        tf = self.text_color if self.highlighted_cells[r][1] is None else self.highlighted_cells[r][1]
                    else:
                        tf = self.text_color
                    if isinstance(self.MT.my_row_index, int):
                        lns = self.MT.data_ref[r][self.MT.my_row_index].split("\n")
                    else:
                        try:
                            lns = self.MT.my_row_index[r].split("\n")
                        except:
                            lns = (f"{r + 1}", )
                    y = fr + self.MT.fl_ins
                    if y + self.MT.half_txt_h > y1:
                        fl = lns[0]
                        t = self.create_text(x, y, text = fl, fill = tf, font = self.MT.my_font, anchor = "w", tag = "t")
                        wd = self.bbox(t)
                        wd = wd[2] - wd[0]
                        if wd > mw:
                            nl = int(len(fl) * (mw / wd)) - 1
                            self.itemconfig(t, text = fl[:nl])
                            wd = self.bbox(t)
                            while wd[2] - wd[0] > mw:
                                nl -= 1
                                self.dchars(t, nl)
                                wd = self.bbox(t)
                    if len(lns) > 1:
                        stl = int((y1 - y) / self.MT.xtra_lines_increment) - 1
                        if stl < 1:
                            stl = 1
                        y += (stl * self.MT.xtra_lines_increment)
                        if y + self.MT.half_txt_h < sr:
                            for i in range(stl, len(lns)):
                                txt = lns[i]
                                t = self.create_text(x, y, text = txt, fill = tf, font = self.MT.my_font, anchor = "w", tag = "t")
                                wd = self.bbox(t)
                                wd = wd[2] - wd[0]
                                if wd > mw:
                                    nl = int(len(txt) * (mw / wd)) - 1
                                    self.itemconfig(t, text = txt[:nl])
                                    wd = self.bbox(t)
                                    while wd[2] - wd[0] > mw:
                                        nl -= 1
                                        self.dchars(t, nl)
                                        wd = self.bbox(t)
                                y += self.MT.xtra_lines_increment
                                if y + self.MT.half_txt_h > sr:
                                    break
            self.create_line(self.current_width - 1, y1, self.current_width - 1, y_stop, fill = self.row_index_border_color, width = 1, tag = "v")
        except:
            return

    def GetCellCoords(self, event = None, r = None, c = None):
        pass


class MainTable(tk.Canvas):
    def __init__(self,
                 parentframe = None,
                 column_width = None,
                 column_headers_canvas = None,
                 row_index_canvas = None,
                 headers = None,
                 header_height = None,
                 row_height = None,
                 data_reference = None,
                 row_index = None,
                 font = None,
                 header_font = None,
                 align = None,
                 width = None,
                 height = None,
                 table_background = "white",
                 grid_color = "gray15",
                 text_color = "black",
                 selected_cells_background = "gray85",
                 selected_cells_foreground = "white",
                 displayed_columns = [],
                 all_columns_displayed = True):
        tk.Canvas.__init__(self,
                           parentframe,
                           width = width,
                           height = height,
                           background = table_background,
                           highlightthickness = 0)
        self.min_rh = 0
        self.hdr_min_rh = 0
        self.extra_motion_func = None
        self.extra_b1_press_func = None
        self.extra_b1_motion_func = None
        self.extra_b1_release_func = None
        self.extra_double_b1_func = None
        self.extra_ctrl_c_func = None
        self.extra_ctrl_x_func = None
        self.extra_ctrl_v_func = None
        self.extra_ctrl_z_func = None
        self.extra_delete_key_func = None
        self.extra_edit_cell_func = None
        self.selection_binding_func = None # function to run when a spreadsheet selection event occurs
        self.drag_selection_binding_func = None
        self.single_selection_enabled = False
        self.drag_selection_enabled = False
        self.multiple_selection_enabled = False
        self.arrowkeys_enabled = False
        self.new_row_width = 0
        self.new_header_height = 0
        self.parentframe = parentframe
        self.row_width_resize_bbox = tuple()
        self.header_height_resize_bbox = tuple()
        self.CH = column_headers_canvas
        self.CH.MT = self
        self.CH.RI = row_index_canvas
        self.RI = row_index_canvas
        self.RI.MT = self
        self.RI.CH = column_headers_canvas
        self.TL = None                # is set from within TopLeftRectangle() __init__
        self.total_rows = 0
        self.total_cols = 0
        self.data_ref = data_reference
        if self.data_ref:
            self.total_rows = len(self.data_ref)
        if self.data_ref and all_columns_displayed:
            self.total_cols = len(max(self.data_ref, key = len))
        elif self.data_ref and not all_columns_displayed:
            self.total_cols = len(displayed_columns)
        self.displayed_columns = displayed_columns
        self.all_columns_displayed = all_columns_displayed

        self.grid_color = grid_color
        self.text_color = text_color
        self.selected_cells_background = selected_cells_background
        self.selected_cells_foreground = selected_cells_foreground
        self.table_background = table_background
        
        self.align = align
        self.my_font = font
        self.fnt_fam = font[0]
        self.fnt_sze = font[1]
        self.fnt_wgt = font[2]
        self.my_hdr_font = header_font
        self.hdr_fnt_fam = header_font[0]
        self.hdr_fnt_sze = header_font[1]
        self.hdr_fnt_wgt = header_font[2]

        self.txt_measure_canvas = tk.Canvas(self)
        self.table_dropdown = None
        self.table_dropdown_id = None
        self.table_dropdown_value = None
        self.default_cw = column_width
        self.default_rh = int(row_height)
        self.default_hh = int(header_height)
        self.set_fnt_help()
        self.set_hdr_fnt_help()
        
        # headers, default is range of numbers
        if headers:
            self.my_hdrs = headers
        else:
            self.my_hdrs = []
                
        # row indexes, default is range of numbers
        if isinstance(row_index, int):
            self.my_row_index = row_index
        else:
            if row_index:
                self.my_row_index = row_index
            else:
                self.my_row_index = []

        self.col_positions = [0]
        self.row_positions = [0]
        self.reset_col_positions()
        self.reset_row_positions()

        self.currently_selected = tuple() # can be a row ("row",row number) or column ("column",column number) or cell (row number,column number)
        self.selected_cells = set()
        self.selected_cols = set()
        self.selected_rows = set()
        self.highlighted_cells = {}

        self.undo_storage = deque(maxlen = 20)

        self.bind("<Motion>", self.mouse_motion)
        self.bind("<Shift-ButtonPress-1>",self.shift_b1_press)
        self.bind("<ButtonPress-1>", self.b1_press)
        self.bind("<B1-Motion>", self.b1_motion)
        self.bind("<ButtonRelease-1>", self.b1_release)
        self.bind("<Double-Button-1>", self.double_b1)
        self.bind("<Configure>", self.refresh)
        self.bind("<MouseWheel>", self.mousewheel)

        self.rc_popup_menu = tk.Menu(self, tearoff = 0)
        self.rc_popup_menu.add_command(label = "Select all (Ctrl-a)",
                                       command = self.select_all)
        self.rc_popup_menu.add_separator()
        self.rc_popup_menu.add_command(label = "Cut (Ctrl-x)",
                                       command = self.ctrl_x)
        self.rc_popup_menu.add_separator()
        self.rc_popup_menu.add_command(label = "Copy (Ctrl-c)",
                                       command = self.ctrl_c)
        self.rc_popup_menu.add_separator()
        self.rc_popup_menu.add_command(label = "Paste (Ctrl-v)",
                                       command = self.ctrl_v)
        self.rc_popup_menu.add_separator()
        self.rc_popup_menu.add_command(label = "Delete (Del)",
                                       command = self.delete_key)

    def refresh(self, event = None):
        self.main_table_redraw_grid_and_text(True, True)

    def shift_b1_press(self, event = None):
        if self.drag_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            rowsel = int(self.identify_row(y = event.y))
            colsel = int(self.identify_col(x = event.x))
            if rowsel < len(self.row_positions) - 1 and colsel < len(self.col_positions) - 1 and (rowsel, colsel) not in self.selected_cells:
                if self.currently_selected and isinstance(self.currently_selected[0], int):
                    min_r = self.currently_selected[0]
                    min_c = self.currently_selected[1]
                    self.selected_cells = set()
                    self.CH.selected_cells = defaultdict(int)
                    self.RI.selected_cells = defaultdict(int)
                    self.selected_cols = set()
                    self.selected_rows = set()
                    if rowsel >= min_r and colsel >= min_c:
                        for r in range(min_r, rowsel + 1):
                            for c in range(min_c, colsel + 1):
                                self.selected_cells.add((r, c))
                                self.RI.selected_cells[r] += 1
                                self.CH.selected_cells[c] += 1
                    elif rowsel >= min_r and min_c >= colsel:
                        for r in range(min_r, rowsel + 1):
                            for c in range(colsel, min_c + 1):
                                self.selected_cells.add((r, c))
                                self.RI.selected_cells[r] += 1
                                self.CH.selected_cells[c] += 1
                    elif min_r >= rowsel and colsel >= min_c:
                        for r in range(rowsel, min_r + 1):
                            for c in range(min_c, colsel + 1):
                                self.selected_cells.add((r, c))
                                self.RI.selected_cells[r] += 1
                                self.CH.selected_cells[c] += 1
                    elif min_r >= rowsel and min_c >= colsel:
                        for r in range(rowsel, min_r + 1):
                            for c in range(colsel, min_c + 1):
                                self.selected_cells.add((r, c))
                                self.RI.selected_cells[r] += 1
                                self.CH.selected_cells[c] += 1
                else:
                    self.select_cell(rowsel, colsel, redraw = False)
                self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
                if self.selection_binding_func is not None:
                    self.selection_binding_func(("cell", ) + tuple(self.currently_selected))

    def basic_bindings(self, onoff = "enable"):
        if onoff == "enable":
            self.bind("<Motion>", self.mouse_motion)
            self.bind("<ButtonPress-1>", self.b1_press)
            self.bind("<B1-Motion>", self.b1_motion)
            self.bind("<ButtonRelease-1>", self.b1_release)
            self.bind("<Double-Button-1>", self.double_b1)
            self.bind("<MouseWheel>", self.mousewheel)
        elif onoff == "disable":
            self.unbind("<Motion>")
            self.unbind("<ButtonPress-1>")
            self.unbind("<B1-Motion>")
            self.unbind("<ButtonRelease-1>")
            self.unbind("<Double-Button-1>")
            self.unbind("<MouseWheel>")

    def edit_bindings(self, onoff = "enable"):
        if onoff.lower() == "enable":
            self.bind("<Control-c>", self.ctrl_c)
            self.bind("<Control-C>", self.ctrl_c)
            self.bind("<Control-x>", self.ctrl_x)
            self.bind("<Control-X>", self.ctrl_x)
            self.bind("<Control-v>", self.ctrl_v)
            self.bind("<Control-V>", self.ctrl_v)
            self.bind("<Control-z>", self.ctrl_z)
            self.bind("<Control-Z>", self.ctrl_z)
            self.bind("<Delete>", self.delete_key)
            if str(get_os()) == "Darwin":
                self.bind("<2>", self.rc)
            else:
                self.bind("<3>", self.rc)
            self.bind_cell_edit(True)
            self.RI.bind("<Control-c>", self.ctrl_c)
            self.RI.bind("<Control-C>", self.ctrl_c)
            self.RI.bind("<Control-x>", self.ctrl_x)
            self.RI.bind("<Control-X>", self.ctrl_x)
            self.RI.bind("<Control-v>", self.ctrl_v)
            self.RI.bind("<Control-V>", self.ctrl_v)
            self.RI.bind("<Control-z>", self.ctrl_z)
            self.RI.bind("<Control-Z>", self.ctrl_z)
            self.RI.bind("<Delete>", self.delete_key)
            self.CH.bind("<Control-c>", self.ctrl_c)
            self.CH.bind("<Control-C>", self.ctrl_c)
            self.CH.bind("<Control-x>", self.ctrl_x)
            self.CH.bind("<Control-X>", self.ctrl_x)
            self.CH.bind("<Control-v>", self.ctrl_v)
            self.CH.bind("<Control-V>", self.ctrl_v)
            self.CH.bind("<Control-z>", self.ctrl_z)
            self.CH.bind("<Control-Z>", self.ctrl_z)
            self.CH.bind("<Delete>", self.delete_key)
        elif onoff.lower() == "disable":
            self.unbind("<Control-c>")
            self.unbind("<Control-C>")
            self.unbind("<Control-x>")
            self.unbind("<Control-X>")
            self.unbind("<Control-v>")
            self.unbind("<Control-V>")
            self.unbind("<Delete>")
            if str(get_os()) == "Darwin":
                self.unbind("<2>")
            else:
                self.unbind("<3>")
            self.bind_cell_edit(False)
            self.RI.unbind("<Control-c>")
            self.RI.unbind("<Control-C>")
            self.RI.unbind("<Control-x>")
            self.RI.unbind("<Control-X>")
            self.RI.unbind("<Control-v>")
            self.RI.unbind("<Control-V>")
            self.RI.unbind("<Delete>")
            self.CH.unbind("<Control-c>")
            self.CH.unbind("<Control-C>")
            self.CH.unbind("<Control-x>")
            self.CH.unbind("<Control-X>")
            self.CH.unbind("<Control-v>")
            self.CH.unbind("<Control-V>")
            self.CH.unbind("<Control-z>")
            self.CH.unbind("<Control-Z>")
            self.CH.unbind("<Delete>")

    def bind_cell_edit(self, enable = True):
        if enable:
            for c in """abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ""":
                self.bind(f"<{c}>", self.edit_cell_)
            for c in "1234567890":
                self.bind(c, self.edit_cell_)
            self.bind("<F2>", self.edit_cell_)
            self.bind("<Double-Button-1>", self.edit_cell_)
        else:
            for c in """abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ""":
                self.unbind(f"<{c}>")
            for c in "1234567890":
                self.unbind(c)
            self.unbind("<F2>")
            self.unbind("<Double-Button-1>")

    def show_ctrl_outline(self, t = "cut", canvas = "table", start_cell = (0, 0), end_cell = (0, 0)):
        if t != "cut":
            self.create_rectangle(self.col_positions[start_cell[0]] + 1,
                                  self.row_positions[start_cell[1]] + 1,
                                  self.col_positions[end_cell[0]],
                                  self.row_positions[end_cell[1]],
                                  fill = "",
                                  width = 4,
                                  outline = self.grid_color,
                                  tag = "ctrl")
        else:
            self.create_rectangle(self.col_positions[start_cell[0]] + 1,
                                  self.row_positions[start_cell[1]] + 1,
                                  self.col_positions[end_cell[0]],
                                  self.row_positions[end_cell[1]],
                                  fill = "",
                                  dash = (10, 5),
                                  width = 4,
                                  outline = self.table_background,
                                  tag = "ctrl")
        self.tag_raise("ctrl")
        self.after(1000, self.del_ctrl_outline)

    def del_ctrl_outline(self, event = None):
        self.delete("ctrl")

    def ctrl_c(self, event = None):
        if not self.anything_selected():
            return
        s = io.StringIO()
        writer = csv_module.writer(s, dialect = csv_module.excel_tab, lineterminator = "\n")
        if self.selected_cols:
            x1 = self.get_min_selected_cell_x()
            x2 = self.get_max_selected_cell_x() + 1
            y1 = 0
            y2 = len(self.row_positions) - 1
        elif self.selected_rows:
            x1 = 0
            x2 = len(self.col_positions) - 1
            y1 = self.get_min_selected_cell_y()
            y2 = self.get_max_selected_cell_y() + 1
        else:
            x1 = self.get_min_selected_cell_x()
            x2 = self.get_max_selected_cell_x() + 1
            y1 = self.get_min_selected_cell_y()
            y2 = self.get_max_selected_cell_y() + 1
        if self.all_columns_displayed:
            for r in range(y1, y2):
                l_ = [self.data_ref[r][c] for c in range(x1, x2)]
                writer.writerow(l_)
        else:
            for r in range(y1, y2):
                l_ = [self.data_ref[r][self.displayed_columns[c]] for c in range(x1, x2)]
                writer.writerow(l_)
        self.clipboard_clear()
        s = s.getvalue().rstrip()
        self.clipboard_append(s)
        self.update()
        if self.extra_ctrl_c_func is not None:
            self.extra_ctrl_c_func()
            
    def ctrl_x(self, event = None):
        if not self.anything_selected():
            return
        undo_storage = {}
        s = io.StringIO()
        writer = csv_module.writer(s, dialect = csv_module.excel_tab, lineterminator = "\n")
        if self.selected_cols:
            x1 = self.get_min_selected_cell_x()
            x2 = self.get_max_selected_cell_x() + 1
            y1 = 0
            y2 = len(self.row_positions) - 1
        elif self.selected_rows:
            x1 = 0
            x2 = len(self.col_positions) - 1
            y1 = self.get_min_selected_cell_y()
            y2 = self.get_max_selected_cell_y() + 1
        else:
            x1 = self.get_min_selected_cell_x()
            x2 = self.get_max_selected_cell_x() + 1
            y1 = self.get_min_selected_cell_y()
            y2 = self.get_max_selected_cell_y() + 1
        undo_storage = {}
        if self.all_columns_displayed:
            for r in range(y1, y2):
                l_ = []
                for c in range(x1, x2):
                    sl = f"{self.data_ref[r][c]}"
                    l_.append(sl)
                    undo_storage[(r, c)] = sl
                    self.data_ref[r][c] = ""
                writer.writerow(l_)
                
        else:
            for r in range(y1, y2):
                l_ = []
                for c in range(x1, x2):
                    sl = f"{self.data_ref[r][self.displayed_columns[c]]}"
                    l_.append(sl)
                    undo_storage[(r, self.displayed_columns[c])] = sl
                    self.data_ref[r][self.displayed_columns[c]] = ""
                writer.writerow(l_)
        self.undo_storage.append(zlib.compress(pickle.dumps(undo_storage)))
        self.clipboard_clear()
        s = s.getvalue().rstrip()
        self.clipboard_append(s)
        self.update()
        self.refresh()
        self.show_ctrl_outline(t = "cut", canvas = "table", start_cell = (x1, y1), end_cell = (x2, y2))
        if self.extra_ctrl_x_func is not None:
            self.extra_ctrl_x_func()

    def ctrl_v(self, event = None):
        if not self.anything_selected():
            return
        try:
            data = self.clipboard_get()
        except:
            return
        nd = []
        for r in csv_module.reader(io.StringIO(data), delimiter = "\t", quotechar = '"', skipinitialspace = True):
            try:
                nd.append(r[:len(r) - next(i for i, c in enumerate(reversed(r)) if c)])
            except:
                continue
        if not nd:
            return
        data = nd
        numcols = len(max(data, key = len))
        numrows = len(data)
        for rn, r in enumerate(data):
            if len(r) < numcols:
                data[rn].extend(list(repeat("", numcols - len(r))))
        undo_storage = {}
        if self.selected_cols:
            x1 = self.get_min_selected_cell_x()
            y1 = 0
        elif self.selected_rows:
            x1 = 0
            y1 = self.get_min_selected_cell_y()
        else:
            x1 = self.get_min_selected_cell_x()
            y1 = self.get_min_selected_cell_y()
        if x1 + numcols > len(self.col_positions) - 1:
            numcols = len(self.col_positions) - 1 - x1
        if y1 + numrows > len(self.row_positions) - 1:
            numrows = len(self.row_positions) - 1 - y1 
        if self.all_columns_displayed:
            for ndr, r in enumerate(range(y1, y1 + numrows)):
                l_ = []
                for ndc, c in enumerate(range(x1, x1 + numcols)):
                    s = f"{self.data_ref[r][c]}"
                    l_.append(s)
                    undo_storage[(r, c)] = s
                    self.data_ref[r][c] = data[ndr][ndc]
        else:
            for ndr, r in enumerate(range(y1, y1 + numrows)):
                l_ = []
                for ndc, c in enumerate(range(x1, x1 + numcols)):
                    s = f"{self.data_ref[r][self.displayed_columns[c]]}"
                    l_.append(s)
                    undo_storage[(r, self.displayed_columns[c])] = s
                    self.data_ref[r][self.displayed_columns[c]] = data[ndr][ndc]
        self.undo_storage.append(zlib.compress(pickle.dumps(undo_storage)))
        self.refresh()
        self.show_ctrl_outline(t = "paste", canvas = "table", start_cell = (x1, y1), end_cell = (x1 + numcols, y1 + numrows))
        if self.extra_ctrl_v_func is not None:
            self.extra_ctrl_v_func()

    def ctrl_z(self, event = None):
        if not self.undo_storage:
            return
        start_row = float("inf")
        start_col = float("inf")
        undo_storage = pickle.loads(zlib.decompress(self.undo_storage.pop()))
        for (r, c), v in undo_storage.items():
            if r < start_row:
                start_row = r
            if c < start_col:
                start_col = c
            self.data_ref[r][c] = v
        self.select_cell(start_row, start_col)
        self.see(r = start_row, c = start_col, keep_yscroll = False, keep_xscroll = False, bottom_right_corner = False, check_cell_visibility = True, redraw = False)
        self.refresh()
        if self.extra_ctrl_z_func is not None:
            self.extra_ctrl_z_func()

    def delete_key(self, event = None):
        if not self.anything_selected():
            return
        undo_storage = {}
        if self.selected_cols:
            x1 = self.get_min_selected_cell_x()
            x2 = self.get_max_selected_cell_x() + 1
            y1 = 0
            y2 = len(self.row_positions) - 1
        elif self.selected_rows:
            x1 = 0
            x2 = len(self.col_positions) - 1
            y1 = self.get_min_selected_cell_y()
            y2 = self.get_max_selected_cell_y() + 1
        else:
            x1 = self.get_min_selected_cell_x()
            x2 = self.get_max_selected_cell_x() + 1
            y1 = self.get_min_selected_cell_y()
            y2 = self.get_max_selected_cell_y() + 1
        if self.all_columns_displayed:
            for r in range(y1, y2):
                for c in range(x1, x2):
                    undo_storage[(r, c)] = f"{self.data_ref[r][c]}"
                    self.data_ref[r][c] = ""
        else:
            for r in range(y1, y2):
                for c in range(x1, x2):
                    undo_storage[(r, self.displayed_columns[c])] = f"{self.data_ref[r][self.displayed_columns[c]]}"
                    self.data_ref[r][self.displayed_columns[c]] = ""
        self.undo_storage.append(zlib.compress(pickle.dumps(undo_storage)))
        self.refresh()
        if self.extra_delete_key_func is not None:
            self.extra_delete_key_func()
            
    def bind_arrowkeys(self, event = None):
        self.arrowkeys_enabled = True
        for canvas in (self, self.CH, self.RI, self.TL):
            canvas.bind("<Up>", self.arrowkey_UP)
            canvas.bind("<Right>", self.arrowkey_RIGHT)
            canvas.bind("<Down>", self.arrowkey_DOWN)
            canvas.bind("<Left>", self.arrowkey_LEFT)
            canvas.bind("<Prior>", self.page_UP)
            canvas.bind("<Next>", self.page_DOWN)

    def unbind_arrowkeys(self, event = None):
        self.arrowkeys_enabled = False
        for canvas in (self, self.CH, self.RI, self.TL):
            canvas.unbind("<Up>")
            canvas.unbind("<Right>")
            canvas.unbind("<Down>")
            canvas.unbind("<Left>")
            canvas.unbind("<Prior>")
            canvas.unbind("<Next>")

    def see(self, r = None, c = None, keep_yscroll = False, keep_xscroll = False, bottom_right_corner = False, check_cell_visibility = True,
            redraw = True):
        if check_cell_visibility:
            visible = self.cell_is_completely_visible(r = r, c = c)
        else:
            visible = False
        if not visible:
            if bottom_right_corner:
                if r is not None and not keep_yscroll:
                    y = self.row_positions[r + 1] + 1 - self.winfo_height()
                    args = ("moveto", y / (self.row_positions[-1] + 100))
                    self.yview(*args)
                    self.RI.yview(*args)
                    if redraw:
                        self.main_table_redraw_grid_and_text(redraw_row_index = True)
                if c is not None and not keep_xscroll:
                    x = self.col_positions[c + 1] + 1 - self.winfo_width()
                    args = ("moveto",x / (self.col_positions[-1] + 150))
                    self.xview(*args)
                    self.CH.xview(*args)
                    if redraw:
                        self.main_table_redraw_grid_and_text(redraw_header = True)
            else:
                if r is not None and not keep_yscroll:
                    args = ("moveto", self.row_positions[r] / (self.row_positions[-1] + 100))
                    self.yview(*args)
                    self.RI.yview(*args)
                    self.main_table_redraw_grid_and_text(redraw_row_index = True)
                if c is not None and not keep_xscroll:
                    args = ("moveto", self.col_positions[c] / (self.col_positions[-1] + 150))
                    self.xview(*args)
                    self.CH.xview(*args)
                    if redraw:
                        self.main_table_redraw_grid_and_text(redraw_header = True)

    def cell_is_completely_visible(self, r = 0, c = 0, cell_coords = None):
        cx1, cy1, cx2, cy2 = self.get_canvas_visible_area()
        if cell_coords is None:
            x1, y1, x2, y2 = self.GetCellCoords(r = r, c = c, sel = True)
        else:
            x1, y1, x2, y2 = cell_coords
        if cx1 > x1 or cy1 > y1 or cx2 < x2 or cy2 < y2:
            return False
        return True

    def cell_is_visible(self,r = 0, c = 0, cell_coords = None):
        cx1, cy1, cx2, cy2 = self.get_canvas_visible_area()
        if cell_coords is None:
            x1, y1, x2, y2 = self.GetCellCoords(r = r, c = c, sel = True)
        else:
            x1, y1, x2, y2 = cell_coords
        if x1 <= cx2 or y1 <= cy2 or x2 >= cx1 or y2 >= cy1:
            return True
        return False

    #               EXTERNAL USE ADD ANY SELECTION
    def select(self, r, c, cell,redraw = True):
        if cell is not None:
            r, c = cell[0], cell[1]
            self.RI.selected_cells[r] += 1
            self.CH.selected_cells[c] += 1
            self.selected_cells.add(cell)
            if cell == self.currently_selected:
                self.currently_selected = cell
        elif r is not None:
            self.RI.selected_cells[r] += 1
            self.selected_rows.add(r)
        elif c is not None:
            self.CH.selected_cells[c] += 1
            self.selected_cols.add(c)
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)

    #               INTERNAL DESELECT EVERYTHING ELSE, SELECT SINGLE CELL
    def select_cell(self, r, c, redraw = False):
        r = int(r)
        c = int(c)
        self.currently_selected = (r, c)
        self.selected_cells = {(r, c)}
        self.CH.selected_cells = defaultdict(int)
        self.CH.selected_cells[c] += 1
        self.RI.selected_cells = defaultdict(int)
        self.RI.selected_cells[r] += 1
        self.selected_cols = set()
        self.selected_rows = set()
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.selection_binding_func is not None:
            self.selection_binding_func(("cell", ) + tuple(self.currently_selected))

    def highlight_cells(self, r = 0, c = 0, cells = tuple(), bg = None, fg = None, redraw = False):
        if bg is None and fg is None:
            return
        if cells:
            self.highlighted_cells = {t: (bg, fg) for t in cells}
        else:
            self.highlighted_cells[(r, c)] = (bg, fg)
        if redraw:
            self.main_table_redraw_grid_and_text()

    #               INTERNAL ADD SELECTION CELL
    def add_selection(self, r, c, redraw = False, run_binding_func = True):
        r = int(r)
        c = int(c)
        self.currently_selected = (r, c)
        if (r, c) not in self.selected_cells:
            self.selected_cells.add((r, c))
            self.CH.selected_cells[c] += 1
            self.RI.selected_cells[r] += 1
        self.selected_cols = set()
        self.selected_rows = set()
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.selection_binding_func is not None and run_binding_func:
            self.selection_binding_func(("cell", ) + tuple(self.currently_selected))

    def select_all(self, redraw = True, run_binding_func = True):
        self.deselect("all")
        if len(self.row_positions) > 1 and len(self.col_positions) > 1:
            self.currently_selected = (0, 0)
            cols = tuple(range(len(self.col_positions) - 1))
            for r in range(len(self.row_positions) - 1):
                for c in cols:
                    self.selected_cells.add((r, c))
                    self.CH.selected_cells[c] += 1
                    self.RI.selected_cells[r] += 1
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        if self.selection_binding_func is not None and run_binding_func:
            self.selection_binding_func("all")

    def page_UP(self, event = None):
        if not self.arrowkeys_enabled:
            return
        height = self.winfo_height()
        top = self.canvasy(0)
        scrollto = top - height
        if scrollto < 0:
            scrollto = 0
        args = ("moveto", scrollto / (self.row_positions[-1] + 100))
        self.yview(*args)
        self.RI.yview(*args)
        self.main_table_redraw_grid_and_text(redraw_row_index = True)

    def page_DOWN(self, event = None):
        if not self.arrowkeys_enabled:
            return
        height = self.winfo_height()
        top = self.canvasy(0)
        scrollto = top + height
        end = self.row_positions[-1]
        if scrollto > end  + 100:
            scrollto = end
        args = ("moveto", scrollto / (end + 100))
        self.yview(*args)
        self.RI.yview(*args)
        self.main_table_redraw_grid_and_text(redraw_row_index = True)
        
    def arrowkey_UP(self, event = None):
        if not self.currently_selected or not self.arrowkeys_enabled:
            return
        if self.currently_selected[0] == "row":
            r = self.currently_selected[1]
            if r != 0 and self.RI.row_selection_enabled:
                if self.cell_is_completely_visible(r = r - 1, c = 0):
                    self.RI.select_row(r - 1, redraw = True)
                else:
                    self.RI.select_row(r - 1)
                    self.see(r - 1, 0, keep_xscroll = True, check_cell_visibility = False)
        elif isinstance(self.currently_selected[0],int):
            r = self.currently_selected[0]
            c = self.currently_selected[1]
            if r == 0 and self.CH.col_selection_enabled:
                if self.cell_is_completely_visible(r = r, c = 0):
                    self.CH.select_col(c, redraw = True)
                else:
                    self.CH.select_col(c)
                    self.see(r, c, keep_xscroll = True, check_cell_visibility = False)
            elif r != 0 and (self.single_selection_enabled or self.multiple_selection_enabled):
                if self.cell_is_completely_visible(r = r - 1, c = c):
                    self.select_cell(r - 1, c, redraw = True)
                else:
                    self.select_cell(r - 1, c)
                    self.see(r - 1, c, keep_xscroll = True, check_cell_visibility = False)
                
    def arrowkey_RIGHT(self, event = None):
        if not self.currently_selected or not self.arrowkeys_enabled:
            return
        if self.currently_selected[0] == "row":
            r = self.currently_selected[1]
            if self.single_selection_enabled or self.multiple_selection_enabled:
                if self.cell_is_completely_visible(r = r, c = 0):
                    self.select_cell(r, 0, redraw = True)
                else:
                    self.select_cell(r, 0)
                    self.see(r, 0, keep_yscroll = True, bottom_right_corner = True, check_cell_visibility = False)
        elif self.currently_selected[0] == "column":
            c = self.currently_selected[1]
            if c < self.total_cols - 1 and self.CH.col_selection_enabled:
                if self.cell_is_completely_visible(r = 0, c = c + 1):
                    self.CH.select_col(c + 1, redraw = True)
                else:
                    self.CH.select_col(c + 1)
                    self.see(0, c + 1, keep_yscroll = True, bottom_right_corner = True, check_cell_visibility = False)
        elif isinstance(self.currently_selected[0], int):
            r = self.currently_selected[0]
            c = self.currently_selected[1]
            if c < self.total_cols - 1 and (self.single_selection_enabled or self.multiple_selection_enabled):
                if self.cell_is_completely_visible(r = r, c = c + 1):
                    self.select_cell(r, c + 1, redraw =True)
                else:
                    self.select_cell(r, c + 1)
                    self.see(r, c + 1, keep_yscroll = True, bottom_right_corner = True, check_cell_visibility = False)

    def arrowkey_DOWN(self, event = None):
        if not self.currently_selected or not self.arrowkeys_enabled:
            return
        if self.currently_selected[0] == "row":
            r = self.currently_selected[1]
            if r < self.total_rows - 1 and self.RI.row_selection_enabled:
                if self.cell_is_completely_visible(r = r + 1, c = 0):
                    self.RI.select_row(r + 1, redraw = True)
                else:
                    self.RI.select_row(r + 1)
                    self.see(r + 1, 0, keep_xscroll = True, bottom_right_corner = True, check_cell_visibility = False)
        elif self.currently_selected[0] == "column":
            c = self.currently_selected[1]
            if self.single_selection_enabled or self.multiple_selection_enabled:
                if self.cell_is_completely_visible(r = 0, c = c):
                    self.select_cell(0, c, redraw = True)
                else:
                    self.select_cell(0, c)
                    self.see(0, c, keep_xscroll = True, bottom_right_corner = True, check_cell_visibility = False)
        elif isinstance(self.currently_selected[0],int):
            r = self.currently_selected[0]
            c = self.currently_selected[1]
            if r < self.total_rows - 1 and (self.single_selection_enabled or self.multiple_selection_enabled):
                if self.cell_is_completely_visible(r = r + 1, c = c):
                    self.select_cell(r + 1, c, redraw = True)
                else:
                    self.select_cell(r + 1, c)
                    self.see(r + 1, c, keep_xscroll = True, bottom_right_corner = True, check_cell_visibility = False)
                    
    def arrowkey_LEFT(self, event = None):
        if not self.currently_selected or not self.arrowkeys_enabled:
            return
        if self.currently_selected[0] == "column":
            c = self.currently_selected[1]
            if c != 0 and self.CH.col_selection_enabled:
                if self.cell_is_completely_visible(r = 0, c = c - 1):
                    self.CH.select_col(c - 1, redraw = True)
                else:
                    self.CH.select_col(c - 1)
                    self.see(0, c - 1, keep_yscroll = True, bottom_right_corner = True, check_cell_visibility = False)
        elif isinstance(self.currently_selected[0], int):
            r = self.currently_selected[0]
            c = self.currently_selected[1]
            if c == 0 and self.RI.row_selection_enabled:
                if self.cell_is_completely_visible(r = r, c = 0):
                    self.RI.select_row(r, redraw = True)
                else:
                    self.RI.select_row(r)
                    self.see(r, c, keep_yscroll = True, check_cell_visibility = False)
            elif c != 0 and (self.single_selection_enabled or self.multiple_selection_enabled):
                if self.cell_is_completely_visible(r = r, c = c - 1):
                    self.select_cell(r, c - 1, redraw = True)
                else:
                    self.select_cell(r, c - 1)
                    self.see(r, c - 1, keep_yscroll = True, check_cell_visibility = False)

    def enable_bindings(self, bindings):
        if isinstance(bindings,(list, tuple)):
            for binding in bindings:
                self.enable_bindings_internal(binding)
        elif isinstance(bindings, str):
            self.enable_bindings_internal(bindings)

    def enable_bindings_internal(self, binding):
        if binding == "single":
            self.single_selection_enabled = True
            self.multiple_selection_enabled = False
        elif binding == "multiple":
            self.multiple_selection_enabled = True
            self.single_selection_enabled = False
        elif binding == "drag_select":
            self.drag_selection_enabled = True
            self.bind("<Control-a>", self.select_all)
            self.bind("<Control-A>", self.select_all)
            self.RI.bind("<Control-a>", self.select_all)
            self.RI.bind("<Control-A>", self.select_all)
            self.CH.bind("<Control-a>", self.select_all)
            self.CH.bind("<Control-A>", self.select_all)
        elif binding == "column_width_resize":
            self.CH.enable_bindings("column_width_resize")
        elif binding == "column_select":
            self.CH.enable_bindings("column_select")
        elif binding == "column_height_resize":
            self.CH.enable_bindings("column_height_resize")
        elif binding == "column_drag_and_drop":
            self.CH.enable_bindings("drag_and_drop")
        elif binding == "double_click_column_resize":
            self.CH.enable_bindings("double_click_column_resize")
        elif binding == "row_height_resize":
            self.RI.enable_bindings("row_height_resize")
        elif binding == "double_click_row_resize":
            self.RI.enable_bindings("double_click_row_resize")
        elif binding == "row_width_resize":
            self.RI.enable_bindings("row_width_resize")
        elif binding == "row_select":
            self.RI.enable_bindings("row_select")
        elif binding == "row_drag_and_drop":
            self.RI.enable_bindings("drag_and_drop")
        elif binding == "arrowkeys":
            self.bind_arrowkeys()
        
    def disable_bindings(self, bindings):
        if isinstance(bindings,(list, tuple)):
            for binding in bindings:
                self.disable_bindings_internal(binding)
        elif isinstance(bindings, str):
            self.disable_bindings_internal(bindings)

    def disable_bindings_internal(self, binding):
        if binding == "single":
            self.single_selection_enabled = False
        elif binding == "multiple":
            self.multiple_selection_enabled = False
        elif binding == "drag_select":
            self.drag_selection_enabled = False
            self.unbind("<Control-a>")
            self.unbind("<Control-A>")
            self.RI.unbind("<Control-a>")
            self.RI.unbind("<Control-A>")
            self.CH.unbind("<Control-a>")
            self.CH.unbind("<Control-A>")
        elif binding == "column_width_resize":
            self.CH.disable_bindings("column_width_resize")
        elif binding == "column_select":
            self.CH.disable_bindings("column_select")
        elif binding == "column_height_resize":
            self.CH.disable_bindings("column_height_resize")
        elif binding == "column_drag_and_drop":
            self.CH.disable_bindings("drag_and_drop")
        elif binding == "double_click_column_resize":
            self.CH.disable_bindings("double_click_column_resize")
        elif binding == "row_height_resize":
            self.RI.disable_bindings("row_height_resize")
        elif binding == "double_click_row_resize":
            self.RI.disable_bindings("double_click_row_resize")
        elif binding == "row_width_resize":
            self.RI.disable_bindings("row_width_resize")
        elif binding == "row_select":
            self.RI.disable_bindings("row_select")
        elif binding == "row_drag_and_drop":
            self.RI.disable_bindings("drag_and_drop")
        elif binding == "arrowkeys":
            self.unbind_arrowkeys()

    def reset_mouse_motion_creations(self, event = None):
        self.config(cursor = "")
        self.RI.config(cursor = "")
        self.CH.config(cursor = "")
        self.RI.rsz_w = None
        self.RI.rsz_h = None
        self.CH.rsz_w = None
        self.CH.rsz_h = None
    
    def mouse_motion(self, event):
        if (
            not self.RI.currently_resizing_height and
            not self.RI.currently_resizing_width and
            not self.CH.currently_resizing_height and
            not self.CH.currently_resizing_width
            ):
            mouse_over_resize = False
            x = self.canvasx(event.x)
            y = self.canvasy(event.y)
            if self.RI.width_resizing_enabled and not mouse_over_resize:
                try:
                    x1, y1, x2, y2 = self.row_width_resize_bbox[0], self.row_width_resize_bbox[1], self.row_width_resize_bbox[2], self.row_width_resize_bbox[3]
                    if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                        self.config(cursor = "sb_h_double_arrow")
                        self.RI.config(cursor = "sb_h_double_arrow")
                        self.RI.rsz_w = True
                        mouse_over_resize = True
                except:
                    pass
            if self.CH.height_resizing_enabled and not mouse_over_resize:
                try:
                    x1, y1, x2, y2 = self.header_height_resize_bbox[0], self.header_height_resize_bbox[1], self.header_height_resize_bbox[2], self.header_height_resize_bbox[3]
                    if x >= x1 and y >= y1 and x <= x2 and y <= y2:
                        self.config(cursor = "sb_v_double_arrow")
                        self.CH.config(cursor = "sb_v_double_arrow")
                        self.CH.rsz_h = True
                        mouse_over_resize = True
                except:
                    pass
            if not mouse_over_resize:
                self.reset_mouse_motion_creations()
        if self.extra_motion_func is not None:
            self.extra_motion_func(event)

    def rc(self, event = None):
        self.focus_set()
        if self.identify_col(x = event.x, allow_end = False) is None or self.identify_row(y = event.y, allow_end = False) is None:
            self.deselect("all")
        elif self.single_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            r = self.identify_row(y = event.y)
            c = self.identify_col(x = event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                cols_selected = self.anything_selected(exclude_rows = True, exclude_cells = True)
                rows_selected = self.anything_selected(exclude_columns = True, exclude_cells = True)
                if rows_selected and not cols_selected:
                    x1 = 0
                    x2 = len(self.col_positions) - 1
                    y1 = self.get_min_selected_cell_y()
                    y2 = self.get_max_selected_cell_y()
                elif cols_selected and not rows_selected:
                    x1 = self.get_min_selected_cell_x()
                    x2 = self.get_max_selected_cell_x()
                    y1 = 0
                    y2 = len(self.row_positions) - 1
                else:
                    x1 = self.get_min_selected_cell_x()
                    x2 = self.get_max_selected_cell_x()
                    y1 = self.get_min_selected_cell_y()
                    y2 = self.get_max_selected_cell_y()
                if all(e is not None for e in (x1, x2, y1, y2)) and r >= y1 and c >= x1 and r <= y2 and c <= x2:
                    self.rc_popup_menu.tk_popup(event.x_root, event.y_root)
                else:
                    self.select_cell(r, c, redraw = True)
                    self.rc_popup_menu.tk_popup(event.x_root, event.y_root)
        elif self.multiple_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            r = self.identify_row(y = event.y)
            c = self.identify_col(x = event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                cols_selected = self.anything_selected(exclude_rows = True, exclude_cells = True)
                rows_selected = self.anything_selected(exclude_columns = True, exclude_cells = True)
                if rows_selected and not cols_selected:
                    x1 = 0
                    x2 = len(self.col_positions) - 1
                    y1 = self.get_min_selected_cell_y()
                    y2 = self.get_max_selected_cell_y()
                elif cols_selected and not rows_selected:
                    x1 = self.get_min_selected_cell_x()
                    x2 = self.get_max_selected_cell_x()
                    y1 = 0
                    y2 = len(self.row_positions) - 1
                else:
                    x1 = self.get_min_selected_cell_x()
                    x2 = self.get_max_selected_cell_x()
                    y1 = self.get_min_selected_cell_y()
                    y2 = self.get_max_selected_cell_y()
                if all(e is not None for e in (x1, x2, y1, y2)) and r >= y1 and c >= x1 and r <= y2 and c <= x2:
                    self.rc_popup_menu.tk_popup(event.x_root, event.y_root)
                else:
                    self.add_selection(r, c, redraw = True)
                    self.rc_popup_menu.tk_popup(event.x_root, event.y_root)

    def b1_press(self, event = None):
        self.focus_set()
        x1, y1, x2, y2 = self.get_canvas_visible_area()
        if self.identify_col(x = event.x, allow_end = False) is None or self.identify_row(y = event.y, allow_end = False) is None:
            self.deselect("all")
        elif self.single_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            r = self.identify_row(y = event.y)
            c = self.identify_col(x = event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.select_cell(r, c, redraw = True)
        elif self.multiple_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            r = self.identify_row(y = event.y)
            c = self.identify_col(x = event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.add_selection(r, c, redraw = True)
        elif self.RI.width_resizing_enabled and self.RI.rsz_h is None and self.RI.rsz_w == True:
            self.RI.currently_resizing_width = True
            self.new_row_width = self.RI.current_width + event.x
            x = self.canvasx(event.x)
            self.create_line(x, y1, x, y2, width = 1, fill = self.RI.resizing_line_color, tag = "rwl")
        elif self.CH.height_resizing_enabled and self.CH.rsz_w is None and self.CH.rsz_h == True:
            self.CH.currently_resizing_height = True
            self.new_header_height = self.CH.current_height + event.y
            y = self.canvasy(event.y)
            self.create_line(x1, y, x2, y, width = 1, fill = self.RI.resizing_line_color, tag = "rhl")
        if self.extra_b1_press_func is not None:
            self.extra_b1_press_func(event)

    def get_max_selected_cell_x(self):
        try:
            return max(self.CH.selected_cells)
        except:
            return None

    def get_max_selected_cell_y(self):
        try:
            return max(self.RI.selected_cells)
        except:
            return None

    def get_min_selected_cell_y(self):
        try:
            return min(self.RI.selected_cells)
        except:
            return None

    def get_min_selected_cell_x(self):
        try:
            return min(self.CH.selected_cells)
        except:
            return None
        
    def b1_motion(self, event):
        x1, y1, x2, y2 = self.get_canvas_visible_area()
        if self.drag_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)): 
            end_row = self.identify_row(y = event.y)
            end_col = self.identify_col(x = event.x)
            if end_row < len(self.row_positions) - 1 and end_col < len(self.col_positions) - 1 and len(self.currently_selected) == 2:
                if isinstance(self.currently_selected[0], int):
                    start_row = self.currently_selected[0]
                    start_col = self.currently_selected[1]
                    self.selected_cols = set()
                    self.selected_rows = set()
                    self.selected_cells = set()
                    self.CH.selected_cells = defaultdict(int)
                    self.RI.selected_cells = defaultdict(int)
                    if end_row >= start_row and end_col >= start_col:
                        for c in range(start_col, end_col + 1):
                            self.CH.selected_cells[c] += 1
                            for r in range(start_row, end_row + 1):
                                self.RI.selected_cells[r] += 1
                                self.selected_cells.add((r, c))
                    elif end_row >= start_row and end_col < start_col:
                        for c in range(end_col,start_col + 1):
                            self.CH.selected_cells[c] += 1
                            for r in range(start_row,end_row + 1):
                                self.RI.selected_cells[r] += 1
                                self.selected_cells.add((r,c))
                    elif end_row < start_row and end_col >= start_col:
                        for c in range(start_col, end_col + 1):
                            self.CH.selected_cells[c] += 1
                            for r in range(end_row, start_row + 1):
                                self.RI.selected_cells[r] += 1
                                self.selected_cells.add((r, c))
                    elif end_row < start_row and end_col < start_col:
                        for c in range(end_col,start_col + 1):
                            self.CH.selected_cells[c] += 1
                            for r in range(end_row,start_row + 1):
                                self.RI.selected_cells[r] += 1
                                self.selected_cells.add((r, c))
                    if self.drag_selection_binding_func is not None:
                        self.drag_selection_binding_func(sorted([start_row, end_row]) + sorted([start_col, end_col]))
            if event.x > self.winfo_width():
                try:
                    self.xview_scroll(1, "units")
                    self.CH.xview_scroll(1, "units")
                except:
                    pass
            elif event.x < 0:
                try:
                    self.xview_scroll(-1, "units")
                    self.CH.xview_scroll(-1, "units")
                except:
                    pass
            if event.y > self.winfo_height():
                try:
                    self.yview_scroll(1, "units")
                    self.RI.yview_scroll(1, "units")
                except:
                    pass
            elif event.y < 0:
                try:
                    self.yview_scroll(-1, "units")
                    self.RI.yview_scroll(-1, "units")
                except:
                    pass
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        elif self.RI.width_resizing_enabled and self.RI.rsz_w is not None and self.RI.currently_resizing_width:
            self.RI.delete("rwl")
            self.delete("rwl")
            if event.x >= 0:
                x = self.canvasx(event.x)
                self.new_row_width = self.RI.current_width + event.x
                self.create_line(x, y1, x, y2, width = 1, fill = self.RI.resizing_line_color, tag = "rwl")
            else:
                x = self.RI.current_width + event.x
                if x < self.min_cw:
                    x = int(self.min_cw)
                self.new_row_width = x
                self.RI.create_line(x, y1, x, y2, width = 1, fill = self.RI.resizing_line_color, tag = "rwl")
        elif self.CH.height_resizing_enabled and self.CH.rsz_h is not None and self.CH.currently_resizing_height:
            self.CH.delete("rhl")
            self.delete("rhl")
            if event.y >= 0:
                y = self.canvasy(event.y)
                self.new_header_height = self.CH.current_height + event.y
                self.create_line(x1, y, x2, y, width = 1, fill = self.RI.resizing_line_color, tag = "rhl")
            else:
                y = self.CH.current_height + event.y
                if y < self.hdr_min_rh:
                    y = int(self.hdr_min_rh)
                self.new_header_height = y
                self.CH.create_line(x1, y, x2, y, width = 1, fill = self.RI.resizing_line_color, tag = "rhl")
        
        if self.extra_b1_motion_func is not None:
            self.extra_b1_motion_func(event)
        
    def b1_release(self, event = None):
        if self.RI.width_resizing_enabled and self.RI.rsz_w is not None and self.RI.currently_resizing_width:
            self.delete("rwl")
            self.RI.delete("rwl")
            self.RI.currently_resizing_width = False
            self.RI.set_width(self.new_row_width, set_TL = True)
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        elif self.CH.height_resizing_enabled and self.CH.rsz_h is not None and self.CH.currently_resizing_height:
            self.delete("rhl")
            self.CH.delete("rhl")
            self.CH.currently_resizing_height = False
            self.CH.set_height(self.new_header_height, set_TL = True)
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        self.RI.rsz_w = None
        self.CH.rsz_h = None
        self.mouse_motion(event)
        if self.extra_b1_release_func is not None:
            self.extra_b1_release_func(event)

    def double_b1(self, event = None):
        self.focus_set()
        x1, y1, x2, y2 = self.get_canvas_visible_area()
        if self.identify_col(x = event.x, allow_end = False) is None or self.identify_row(y = event.y, allow_end = False) is None:
            self.deselect("all")
        elif self.single_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            r = self.identify_row(y = event.y)
            c = self.identify_col(x = event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.select_cell(r, c, redraw = True)
        elif self.multiple_selection_enabled and all(v is None for v in (self.RI.rsz_h, self.RI.rsz_w, self.CH.rsz_h, self.CH.rsz_w)):
            r = self.identify_row(y = event.y)
            c = self.identify_col(x = event.x)
            if r < len(self.row_positions) - 1 and c < len(self.col_positions) - 1:
                self.add_selection(r, c, redraw = True)
        if self.extra_double_b1_func is not None:
            self.extra_double_b1_func(event)

    def deselect(self, r = None, c = None, cell = None, redraw = True):
        if r == "all":
            self.selected_cells = set()
            self.currently_selected = tuple()
            self.CH.selected_cells = defaultdict(int)
            self.RI.selected_cells = defaultdict(int)
            self.selected_cols = set()
            self.selected_rows = set()
        if r != "all" and r is not None:
            try: del self.RI.selected_cells[r]
            except: pass
            self.selected_rows.discard(r)
        if c is not None:
            try: del self.CH.selected_cells[c]
            except: pass
            self.selected_cols.discard(c)
        if cell is not None:
            r, c = cell[0], cell[1]
            self.selected_cells.discard(cell)
            self.RI.selected_cells[r] -= 1
            if self.RI.selected_cells[r] < 1:
                del self.RI.selected_cells[r]
            self.CH.selected_cells[c] -= 1
            if self.CH.selected_cells[c] < 1:
                del self.CH.selected_cells[c]
            if cell == self.currently_selected:
                self.currently_selected = tuple()
        if redraw:
            self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)

    def identify_row(self, event = None, y = None, allow_end = True):
        if event is None:
            y2 = self.canvasy(y)
        elif y is None:
            y2 = self.canvasy(event.y)
        r = bisect.bisect_left(self.row_positions, y2)
        if r != 0:
            r -= 1
        if not allow_end:
            if r >= len(self.row_positions) - 1:
                return None
        return r

    def identify_col(self, event = None, x = None, allow_end = True):
        if event is None:
            x2 = self.canvasx(x)
        elif x is None:
            x2 = self.canvasx(event.x)
        c = bisect.bisect_left(self.col_positions, x2)
        if c != 0:
            c -= 1
        if not allow_end:
            if c >= len(self.col_positions) - 1:
                return None
        return c

    def GetCellCoords(self, event = None, r = None, c = None, sel = False):
        # event takes priority as parameter
        if event is not None:
            r = self.identify_row(event)
            c = self.identify_col(event)
        elif r is not None and c is not None:
            if sel:
                return self.col_positions[c] + 1,self.row_positions[r] + 1, self.col_positions[c + 1], self.row_positions[r + 1]
            else:
                return self.col_positions[c], self.row_positions[r], self.col_positions[c + 1], self.row_positions[r + 1]

    def set_xviews(self, *args):
        self.xview(*args)
        self.CH.xview(*args)
        self.main_table_redraw_grid_and_text(redraw_header = True)

    def set_yviews(self, *args):
        self.yview(*args)
        self.RI.yview(*args)
        self.main_table_redraw_grid_and_text(redraw_row_index = True)

    def set_view(self, x_args, y_args):
        self.xview(*x_args)
        self.CH.xview(*x_args)
        self.yview(*y_args)
        self.RI.yview(*y_args)
        self.main_table_redraw_grid_and_text(redraw_row_index = True, redraw_header = True)

    def mousewheel(self, event = None):
        if event.num == 5 or event.delta == -120:
            self.yview_scroll(1, "units")
            self.RI.yview_scroll(1, "units")
        if event.num == 4 or event.delta == 120:
            if self.canvasy(0) <= 0:
                return
            self.yview_scroll(-1, "units")
            self.RI.yview_scroll(-1, "units")
        self.main_table_redraw_grid_and_text(redraw_row_index = True)

    def GetWidthChars(self, width):
        char_w = self.GetTextWidth("_")
        return int(width / char_w)

    def GetTextWidth(self, txt):
        x = self.txt_measure_canvas.create_text(0, 0, text = txt, font = self.my_font)
        b = self.txt_measure_canvas.bbox(x)
        self.txt_measure_canvas.delete(x)
        return b[2] - b[0]

    def GetTextHeight(self, txt):
        x = self.txt_measure_canvas.create_text(0, 0, text = txt, font = self.my_font)
        b = self.txt_measure_canvas.bbox(x)
        self.txt_measure_canvas.delete(x)
        return b[3] - b[1]

    def GetHdrTextWidth(self, txt):
        x = self.txt_measure_canvas.create_text(0, 0, text = txt, font = self.my_hdr_font)
        b = self.txt_measure_canvas.bbox(x)
        self.txt_measure_canvas.delete(x)
        return b[2] - b[0]

    def GetHdrTextHeight(self, txt):
        x = self.txt_measure_canvas.create_text(0, 0, text = txt, font = self.my_hdr_font)
        b = self.txt_measure_canvas.bbox(x)
        self.txt_measure_canvas.delete(x)
        return b[3] - b[1]

    def set_min_cw(self):
        w1 = self.GetHdrTextWidth("XXXX")
        w2 = self.GetTextWidth("XXXX")
        if w1 >= w2:
            self.min_cw = w1
        else:
            self.min_cw = w2
        if self.min_cw > self.CH.max_colwidth:
            self.CH.max_colwidth = self.min_cw * 2
        if self.min_cw > self.default_cw:
            self.default_cw = self.min_cw * 2
                 
    #               QUERY OR SET FONT ("FONT",SIZE)
    def font(self, newfont = None):
        if newfont:
            if (
                not isinstance(newfont, tuple) or
                not isinstance(newfont[0], str) or
                not isinstance(newfont[1], int)
                ):
                raise ValueError("Parameter must be tuple e.g. ('Arial',12,'normal')")
            if len(newfont) > 2:
                raise ValueError("Parameter must be three-tuple")
            else:
                self.my_font = newfont
            self.fnt_fam = newfont[0]
            self.fnt_sze = newfont[1]
            self.fnt_wgt = newfont[2]
            self.set_fnt_help()
        else:
            return self.my_font

    def set_fnt_help(self):
        self.txt_h = self.GetTextHeight("|ZX*'^")
        self.half_txt_h = ceil(self.txt_h / 2)
        self.fl_ins = self.half_txt_h + 3
        self.xtra_lines_increment = int(self.txt_h)
        self.min_rh = self.txt_h + 6
        if self.min_rh < 12:
            self.min_rh = 12
        self.set_min_cw()
        
    #               QUERY OR SET HEADER FONT ("FONT",SIZE)
    def header_font(self, newfont = None):
        if newfont:
            if (
                not isinstance(newfont, tuple) or
                not isinstance(newfont[0], str) or
                not isinstance(newfont[1], int)
                ):
                raise ValueError("Parameter must be tuple e.g. ('Arial',12,'bold')")
            if len(newfont) == 3:
                if not isinstance(newfont[2], str):
                    raise ValueError("Parameter must be tuple e.g. ('Arial',12,'bold')")
            if len(newfont) > 3:
                raise ValueError("Parameter must be three tuple")
            else:
                self.my_hdr_font = newfont
            self.hdr_fnt_fam = newfont[0]
            self.hdr_fnt_sze = newfont[1]
            self.hdr_fnt_wgt = newfont[2]
            self.set_hdr_fnt_help()
        else:
            return self.my_hdr_font

    def set_hdr_fnt_help(self):
        self.hdr_txt_h = self.GetHdrTextHeight("|ZX*'^")
        self.hdr_half_txt_h = ceil(self.hdr_txt_h / 2)
        self.hdr_fl_ins = self.hdr_half_txt_h + 5
        self.hdr_xtra_lines_increment = self.hdr_txt_h
        self.hdr_min_rh = self.hdr_txt_h + 10
        self.set_min_cw()
        self.CH.set_height(self.GetHdrLinesHeight(self.default_hh))

    #               QUERY OR SET DATA REFERENCE
    def data_reference(self, newdataref = None, total_cols = None, total_rows = None, reset_col_positions = True, reset_row_positions = True, redraw = False):
        if isinstance(newdataref, (list, tuple)):
            self.data_ref = newdataref
            self.undo_storage = deque(maxlen = 20)
            if total_cols is None and self.all_columns_displayed:
                try:
                    self.total_cols = len(max(newdataref, key = len))
                except:
                    self.total_cols = 0
            elif isinstance(total_cols, int) and self.all_columns_displayed:
                self.total_cols = total_cols
            if total_rows is not None:
                self.total_rows = total_rows
            else:
                self.total_rows = len(newdataref)
            # total cols parameter here is ignored if not all_columns_displayed
            if reset_col_positions:
                self.reset_col_positions()
            if reset_row_positions:
                self.reset_row_positions()
            if redraw:
                self.main_table_redraw_grid_and_text(redraw_header = True, redraw_row_index = True)
        else:
            return id(self.data_ref)

    def reset_col_positions(self):
        colpos = int(self.default_cw)
        self.col_positions = [0] + list(accumulate(colpos for c in range(self.total_cols)))

    def del_col_position(self, idx, deselect_all = False, preserve_other_selections = False):
        # WORK NEEDED FOR PRESERVE SELECTIONS ?
        if deselect_all:
            self.deselect("all", redraw = False)
        if idx == "end" or len(self.col_positions) <= idx + 1:
            del self.col_positions[-1]
        else:
            w = self.col_positions[idx + 1] - self.col_positions[idx]
            idx += 1
            del self.col_positions[idx]
            self.col_positions[idx:] = [e - w for e in islice(self.col_positions, idx, len(self.col_positions))]
        self.total_cols -= 1

    def insert_col_position(self, idx, width = None, deselect_all = False, preserve_other_selections = False):
        # WORK NEEDED FOR PRESERVE SELECTIONS ?
        if deselect_all:
            self.deselect("all", redraw = False)
        if width is None:
            w = self.default_cw
        else:
            w = width
        if idx == "end" or len(self.col_positions) == idx + 1:
            self.col_positions.append(self.col_positions[-1] + w)
        else:
            idx += 1
            self.col_positions.insert(idx,self.col_positions[idx - 1] + w)
            idx += 1
            self.col_positions[idx:] = [e + w for e in islice(self.col_positions, idx, len(self.col_positions))]
        self.total_cols += 1

    def reset_row_positions(self):
        rowpos = self.GetLinesHeight(self.default_rh)
        self.row_positions = [0] + list(accumulate(rowpos for r in range(self.total_rows)))

    def del_row_position(self, idx, deselect_all = False, preserve_other_selections = False):
        # WORK NEEDED FOR PRESERVE SELECTIONS ?
        if deselect_all:
            self.deselect("all", redraw = False)
        if idx == "end" or len(self.row_positions) <= idx + 1:
            del self.row_positions[-1]
        else:
            w = self.row_positions[idx + 1] - self.row_positions[idx]
            idx += 1
            del self.row_positions[idx]
            self.row_positions[idx:] = [e - w for e in islice(self.row_positions, idx, len(self.row_positions))]
        self.total_rows -= 1

    def insert_row_position(self, idx, height = None, deselect_all = False, preserve_other_selections = False):
        # WORK NEEDED FOR PRESERVE SELECTIONS ?
        #here
        if deselect_all:
            self.deselect("all", redraw = False)
        if height is None:
            h = self.GetLinesHeight(self.default_rh)
        else:
            h = height
        if idx == "end" or len(self.row_positions) == idx + 1:
            self.row_positions.append(self.row_positions[-1] + h)
        else:
            idx += 1
            self.row_positions.insert(idx, self.row_positions[idx - 1] + h)
            idx += 1
            self.row_positions[idx:] = [e + h for e in islice(self.row_positions, idx, len(self.row_positions))]
        self.total_rows += 1

    def move_row_position(self, idx1, idx2):
        if not len(self.row_positions) <= 2:
            if idx1 < idx2:
                height = self.row_positions[idx1 + 1] - self.row_positions[idx1]
                self.row_positions.insert(idx2 + 1, self.row_positions.pop(idx1 + 1))
                for i in range(idx1 + 1, idx2 + 1):
                    self.row_positions[i] -= height
                self.row_positions[idx2 + 1] = self.row_positions[idx2] + height
            else:
                height = self.row_positions[idx1 + 1] - self.row_positions[idx1]
                self.row_positions.insert(idx2 + 1, self.row_positions.pop(idx1 + 1))
                for i in range(idx2 + 2, idx1 + 2):
                    self.row_positions[i] += height
                self.row_positions[idx2 + 1] = self.row_positions[idx2] + height

    def move_col_position(self, idx1, idx2):
        if not len(self.col_positions) <= 2:
            if idx1 < idx2:
                width = self.col_positions[idx1 + 1] - self.col_positions[idx1]
                self.col_positions.insert(idx2 + 1, self.col_positions.pop(idx1 + 1))
                for i in range(idx1 + 1, idx2 + 1):
                    self.col_positions[i] -= width
                self.col_positions[idx2 + 1] = self.col_positions[idx2] + width
            else:
                width = self.col_positions[idx1 + 1] - self.col_positions[idx1]
                self.col_positions.insert(idx2 + 1, self.col_positions.pop(idx1 + 1))
                for i in range(idx2 + 2, idx1 + 2):
                    self.col_positions[i] += width
                self.col_positions[idx2 + 1] = self.col_positions[idx2] + width

    def GetLinesHeight(self, n):
        y = int(self.fl_ins)
        if n == 1:
            y += 6
        else:
            for i in range(n):
                y += self.xtra_lines_increment
        if y < self.min_rh:
            y = int(self.min_rh)
        return y

    def GetHdrLinesHeight(self, n):
        y = int(self.hdr_fl_ins)
        if n == 1:
            y + 10
        else:
            for i in range(n):
                y += self.hdr_xtra_lines_increment
        if y < self.hdr_min_rh:
            y = int(self.hdr_min_rh)
        return y

    #               QUERY OR SET DISPLAYED COLUMNS
    def display_columns(self, indexes = None, enable = None, reset_col_positions = True, number_of_columns = None, deselect_all = True):
        if deselect_all:
            self.deselect("all")
        if indexes is None and enable is None:
            return tuple(self.displayed_columns)
        if indexes is not None:
            self.displayed_columns = indexes
            if enable is None:
                if not self.all_columns_displayed:
                    self.total_cols = len(self.displayed_columns)
        if enable == True:
            self.all_columns_displayed = False
            self.total_cols = len(self.displayed_columns)
        if enable == False:
            self.all_columns_displayed = True
            if number_of_columns is not None:
                self.total_cols = number_of_columns
            else:
                self.total_cols = len(max(self.data_ref, key = len))
        if reset_col_positions:
            self.reset_col_positions()
                
    #               QUERY OR SET HEADER OR HEADERS
    def headers(self, newheaders = None, index = None):
        if newheaders is not None:
            if isinstance(newheaders, (list, tuple)):
                if (
                    isinstance(newheaders,(str, int, float)) and
                    isinstance(index, int)
                    ):
                    self.my_hdrs[index] = f"{newheaders}"
                else:
                    self.my_hdrs = newheaders
                    if self.all_columns_displayed:
                        self.total_cols = len(self.my_hdrs)
            elif isinstance(newheaders, int):
                self.my_hdrs = newheaders
        else:
            if index is not None:
                if isinstance(index, int):
                    return self.my_hdrs[index]
            else:
                return self.my_hdrs

    #               QUERY OR SET ROW INDEX OR ROW INDEXES
    def row_index(self, newindex = None, index = None):
        if isinstance(newindex, int):
            self.my_row_index = newindex
        else:
            if isinstance(newindex, (list, tuple)):
                self.my_row_index = newindex
            else:
                if index is not None:
                    if isinstance(index, int):
                        return self.my_row_index[index]
                else:
                    return self.my_row_index

    #               RETURN TUPLE OF CANVAS VISIBLE COORDS
    def get_canvas_visible_area(self):
        return self.canvasx(0), self.canvasy(0), self.canvasx(self.winfo_width()), self.canvasy(self.winfo_height())

    #               RETURN TUPLE OF VISIBLE ROW INDEXES
    def get_visible_rows(self, y1, y2):
        start_row = bisect.bisect_left(self.row_positions, y1)
        end_row = bisect.bisect_right(self.row_positions, y2)
        if not y2 >= self.row_positions[-1]:
            end_row += 1
        return start_row, end_row

    #               RETURN TUPLE OF VISIBLE COLUMN INDEXES
    def get_visible_columns(self, x1, x2):
        start_col = bisect.bisect_left(self.col_positions, x1)
        end_col = bisect.bisect_right(self.col_positions, x2)
        if not x2 >= self.col_positions[-1]:
            end_col += 1
        return start_col, end_col

    #               USE OF FUNCTIONS HAS BEEN MOSTLY AVOIDED HERE TO INCREASE SPEED AND AVOID REPEATED INDEX LOOKUPS
    #               IF REFRESH SPEED IS LIGHTNING QUICK THERE'S NO NEED FOR DRAWING TEXT IN SPECIFIC CELLS
    def main_table_redraw_grid_and_text(self, redraw_header = False, redraw_row_index = False):
        try:
            last_col_line_pos = self.col_positions[-1] + 1
            last_row_line_pos = self.row_positions[-1] + 1
            self.configure(scrollregion=(0, 0, last_col_line_pos + 150, last_row_line_pos + 100))
            self.delete("all")
            x1 = self.canvasx(0)
            y1 = self.canvasy(0)
            x2 = self.canvasx(self.winfo_width())
            y2 = self.canvasy(self.winfo_height())
            self.row_width_resize_bbox = (x1, y1, x1 + 5, y2)
            self.header_height_resize_bbox = (x1 + 6, y1, x2, y1 + 3)
            start_row = bisect.bisect_left(self.row_positions, y1)
            end_row = bisect.bisect_right(self.row_positions, y2)
            if not y2 >= self.row_positions[-1]:
                end_row += 1
            start_col = bisect.bisect_left(self.col_positions, x1)
            end_col = bisect.bisect_right(self.col_positions, x2)
            if not x2 >= self.col_positions[-1]:
                end_col += 1
            if last_col_line_pos > x2:
                x_stop = x2
            else:
                x_stop = last_col_line_pos
            if last_row_line_pos > y2:
                y_stop = y2
            else:
                y_stop = last_row_line_pos
            cr_ = self.create_rectangle
            ct_ = self.create_text
            sb = y2 + 2
            if start_row > 0:
                selsr = start_row - 1
            else:
                selsr = start_row
            if start_col > 0:
                selsc = start_col - 1
            else:
                selsc = start_col
            if not self.selected_cells:
                if self.selected_rows:
                    for r in range(selsr, end_row - 1):
                        fr = self.row_positions[r]
                        sr = self.row_positions[r + 1]
                        if sr > sb:
                            sr = sb
                        if r in self.selected_rows:
                            cr_(x1, fr + 1, x_stop, sr, fill = self.selected_cells_background, outline = "")
                elif self.selected_cols:
                    for c in range(selsc, end_col - 1):
                        fc = self.col_positions[c]
                        sc = self.col_positions[c + 1]
                        if sc > x2 + 2:
                            sc = x2 + 2
                        if c in self.selected_cols:
                            cr_(fc + 1, y1, sc, y_stop, fill = self.selected_cells_background, outline = "")
            for r in range(start_row - 1, end_row):
                y = self.row_positions[r]
                self.create_line(x1, y, x_stop, y, fill= self.grid_color, width = 1)
            for c in range(start_col - 1, end_col):
                x = self.col_positions[c]
                self.create_line(x, y1, x, y_stop, fill = self.grid_color, width = 1)
            if start_row > 0:
                start_row -= 1
            if start_col > 0:
                start_col -= 1   
            end_row -= 1
            c_2 = self.selected_cells_background if self.selected_cells_background.startswith("#") else Color_Map_[self.selected_cells_background]
            c_2_ = (int(c_2[1:3], 16), int(c_2[3:5], 16), int(c_2[5:], 16))
            rows_ = tuple(range(start_row, end_row))
            if self.all_columns_displayed:
                if self.align == "w":
                    for c in range(start_col, end_col - 1):
                        fc = self.col_positions[c]
                        sc = self.col_positions[c + 1]
                        x = fc + 5
                        mw = sc - fc - 5
                        for r in rows_:
                            fr = self.row_positions[r]
                            sr = self.row_positions[r + 1]
                            if sr > sb:
                                sr = sb
                            if (r, c) in self.highlighted_cells and ((r, c) in self.selected_cells or r in self.selected_rows or c in self.selected_cols):
                                c_1 = self.highlighted_cells[(r, c)][0] if self.highlighted_cells[(r, c)][0].startswith("#") else Color_Map_[self.highlighted_cells[(r, c)][0]]
                                cr_(fc + 1,
                                    fr + 1,
                                    sc,
                                    sr,
                                    fill = (f"#{int((int(c_1[1:3], 16) + c_2_[0]) / 2):02X}" +
                                            f"{int((int(c_1[3:5], 16) + c_2_[1]) / 2):02X}" +
                                            f"{int((int(c_1[5:], 16) + c_2_[2]) / 2):02X}"),
                                    outline = "")
                                tf = self.selected_cells_foreground if self.highlighted_cells[(r, c)][1] is None else self.highlighted_cells[(r, c)][1]
                            elif (r, c) in self.selected_cells:
                                cr_(fc + 1, fr + 1, sc, sr, fill = self.selected_cells_background, outline = "")
                                tf = self.selected_cells_foreground
                            elif (r, c) in self.highlighted_cells and r not in self.selected_rows and c not in self.selected_cols:
                                cr_(fc + 1, fr + 1, sc, sr, fill = self.highlighted_cells[(r, c)][0], outline = "")
                                tf = self.text_color if self.highlighted_cells[(r, c)][1] is None else self.highlighted_cells[(r, c)][1]
                            elif r in self.selected_rows or c in self.selected_cols:
                                tf = self.selected_cells_foreground
                            else:
                                tf = self.text_color
                            if x > x2:
                                continue
                            try:
                                lns = self.data_ref[r][c]
                                if isinstance(lns, str):
                                    lns = lns.split("\n")
                                else:
                                    lns = (f"{lns}",)
                                y = fr + self.fl_ins
                                if y + self.half_txt_h > y1:
                                    fl = lns[0]
                                    t = ct_(x, y, text = fl, fill = tf, font = self.my_font, anchor = "w")
                                    wd = self.bbox(t)
                                    wd = wd[2] - wd[0]
                                    if wd > mw:
                                        #nl = int(mw * (len(fl) / wd)) - 1
                                        nl = int(len(fl) * (mw / wd)) - 1
                                        self.itemconfig(t,text=fl[:nl])
                                        wd = self.bbox(t)
                                        while wd[2] - wd[0] > mw:
                                            nl -= 1
                                            self.dchars(t,nl)
                                            wd = self.bbox(t)
                                if len(lns) > 1:
                                    stl = int((y1 - y) / self.xtra_lines_increment) - 1
                                    if stl < 1:
                                        stl = 1
                                    y += (stl * self.xtra_lines_increment)
                                    if y + self.half_txt_h < sr:
                                        for i in range(stl, len(lns)):
                                            txt = lns[i]
                                            t = ct_(x, y, text = txt, fill = tf, font = self.my_font, anchor = "w")
                                            wd = self.bbox(t)
                                            wd = wd[2] - wd[0]
                                            if wd > mw:
                                                #nl = int(mw * (len(txt) / wd)) - 1
                                                nl = int(len(txt) * (mw / wd)) - 1
                                                self.itemconfig(t,text=txt[:nl])
                                                wd = self.bbox(t)
                                                while wd[2] - wd[0] > mw:
                                                    nl -= 1
                                                    self.dchars(t,nl)
                                                    wd = self.bbox(t)
                                            y += self.xtra_lines_increment
                                            if y + self.half_txt_h > sr:
                                                break
                            except:
                                continue
                elif self.align == "center":
                    for c in range(start_col, end_col - 1):
                        fc = self.col_positions[c]
                        stop = fc + 5
                        sc = self.col_positions[c + 1]
                        mw = sc - fc - 5
                        x = fc + floor(mw / 2)
                        for r in rows_:
                            fr = self.row_positions[r]
                            sr = self.row_positions[r + 1]
                            if sr > sb:
                                sr = sb
                            if (r, c) in self.highlighted_cells and ((r, c) in self.selected_cells or r in self.selected_rows or c in self.selected_cols):
                                c_1 = self.highlighted_cells[(r, c)][0] if self.highlighted_cells[(r, c)][0].startswith("#") else Color_Map_[self.highlighted_cells[(r, c)][0]]
                                cr_(fc + 1,
                                    fr + 1,
                                    sc,
                                    sr,
                                    fill = (f"#{int((int(c_1[1:3], 16) + c_2_[0]) / 2):02X}" +
                                            f"{int((int(c_1[3:5], 16) + c_2_[1]) / 2):02X}" +
                                            f"{int((int(c_1[5:], 16) + c_2_[2]) / 2):02X}"),
                                    outline = "")
                                tf = self.selected_cells_foreground if self.highlighted_cells[(r, c)][1] is None else self.highlighted_cells[(r, c)][1]
                            elif (r, c) in self.selected_cells:
                                cr_(fc + 1, fr + 1, sc, sr, fill = self.selected_cells_background, outline = "")
                                tf = self.selected_cells_foreground
                            elif (r, c) in self.highlighted_cells and r not in self.selected_rows and c not in self.selected_cols:
                                cr_(fc + 1, fr + 1, sc, sr, fill = self.highlighted_cells[(r, c)][0], outline = "")
                                tf = self.text_color if self.highlighted_cells[(r, c)][1] is None else self.highlighted_cells[(r, c)][1]
                            elif r in self.selected_rows or c in self.selected_cols:
                                tf = self.selected_cells_foreground
                            else:
                                tf = self.text_color
                            if stop > x2:
                                continue
                            try:
                                lns = self.data_ref[r][c]
                                if isinstance(lns, str):
                                    lns = lns.split("\n")
                                else:
                                    lns = (f"{lns}", )
                                fl = lns[0]
                                y = fr + self.fl_ins
                                if y + self.half_txt_h > y1:
                                    t = ct_(x, y, text = fl, fill = tf, font = self.my_font, anchor = "center")
                                    wd = self.bbox(t)
                                    wd = wd[2] - wd[0]
                                    if wd > mw:
                                        tl = len(fl)
                                        slce = tl - floor(tl * (mw / wd))
                                        if slce % 2:
                                            slce += 1
                                        else:
                                            slce += 2
                                        slce = int(slce / 2)
                                        fl = fl[slce:tl - slce]
                                        self.itemconfig(t, text = fl)
                                        wd = self.bbox(t)
                                        while wd[2] - wd[0] > mw:
                                            fl = fl[1: - 1]
                                            self.itemconfig(t, text = fl)
                                            wd = self.bbox(t)
                                if len(lns) > 1:
                                    stl = int((y1 - y) / self.xtra_lines_increment) - 1
                                    if stl < 1:
                                        stl = 1
                                    y += (stl * self.xtra_lines_increment)
                                    if y + self.half_txt_h < sr:
                                        for i in range(stl,len(lns)):
                                            txt = lns[i]
                                            t = ct_(x, y, text = txt, fill = tf, font = self.my_font, anchor = "center")
                                            wd = self.bbox(t)
                                            wd = wd[2] - wd[0]
                                            if wd > mw:
                                                tl = len(txt)
                                                slce = tl - floor(tl * (mw / wd))
                                                if slce % 2:
                                                    slce += 1
                                                else:
                                                    slce += 2
                                                slce = int(slce / 2)
                                                txt = txt[slce:tl - slce]
                                                self.itemconfig(t, text = txt)
                                                wd = self.bbox(t)
                                                while wd[2] - wd[0] > mw:
                                                    txt = txt[1: - 1]
                                                    self.itemconfig(t, text = txt)
                                                    wd = self.bbox(t)
                                            y += self.xtra_lines_increment
                                            if y + self.half_txt_h > sr:
                                                break
                            except:
                                continue
            else:
                if self.align == "w":
                    for c in range(start_col, end_col - 1):
                        fc = self.col_positions[c]
                        sc = self.col_positions[c + 1]
                        x = fc + 5
                        mw = sc - fc - 5
                        for r in rows_:
                            fr = self.row_positions[r]
                            sr = self.row_positions[r + 1]
                            if sr > sb:
                                sr = sb
                            if (r, self.displayed_columns[c]) in self.highlighted_cells and ((r, c) in self.selected_cells or r in self.selected_rows or c in self.selected_cols):
                                c_1 = self.highlighted_cells[(r, self.displayed_columns[c])][0] if self.highlighted_cells[(r, self.displayed_columns[c])][0].startswith("#") else Color_Map_[self.highlighted_cells[(r, self.displayed_columns[c])][0]]
                                cr_(fc + 1,
                                    fr + 1,
                                    sc,
                                    sr,
                                    fill = (f"#{int((int(c_1[1:3], 16) + c_2_[0]) / 2):02X}" +
                                            f"{int((int(c_1[3:5], 16) + c_2_[1]) / 2):02X}" +
                                            f"{int((int(c_1[5:], 16) + c_2_[2]) / 2):02X}"),
                                    outline = "")
                                tf = self.selected_cells_foreground if self.highlighted_cells[(r, self.displayed_columns[c])][1] is None else self.highlighted_cells[(r, self.displayed_columns[c])][1]
                            elif (r, c) in self.selected_cells:
                                cr_(fc + 1, fr + 1, sc, sr, fill = self.selected_cells_background, outline = "")
                                tf = self.selected_cells_foreground
                            elif (r, self.displayed_columns[c]) in self.highlighted_cells and r not in self.selected_rows and c not in self.selected_cols:
                                cr_(fc + 1, fr + 1, sc, sr, fill = self.highlighted_cells[(r, self.displayed_columns[c])][0], outline = "")
                                tf = self.text_color if self.highlighted_cells[(r, self.displayed_columns[c])][1] is None else self.highlighted_cells[(r, self.displayed_columns[c])][1]
                            elif r in self.selected_rows or c in self.selected_cols:
                                tf = self.selected_cells_foreground
                            else:
                                tf = self.text_color
                            if x > x2:
                                continue
                            try:
                                lns = self.data_ref[r][self.displayed_columns[c]]
                                if isinstance(lns, str):
                                    lns = lns.split("\n")
                                else:
                                    lns = (f"{lns}", )
                                y = fr + self.fl_ins
                                if y + self.half_txt_h > y1:
                                    fl = lns[0]
                                    t = ct_(x, y, text = fl, fill = tf, font = self.my_font, anchor = "w")
                                    wd = self.bbox(t)
                                    wd = wd[2] - wd[0]
                                    if wd > mw:
                                        nl = int(len(fl) * (mw / wd)) - 1
                                        self.itemconfig(t, text = fl[:nl])
                                        wd = self.bbox(t)
                                        while wd[2] - wd[0] > mw:
                                            nl -= 1
                                            self.dchars(t, nl)
                                            wd = self.bbox(t)
                                if len(lns) > 1:
                                    stl = int((y1 - y) / self.xtra_lines_increment) - 1
                                    if stl < 1:
                                        stl = 1
                                    y += (stl * self.xtra_lines_increment)
                                    if y + self.half_txt_h < sr:
                                        for i in range(stl, len(lns)):
                                            txt = lns[i]
                                            t = ct_(x, y, text = txt, fill = tf, font = self.my_font, anchor = "w")
                                            wd = self.bbox(t)
                                            wd = wd[2] - wd[0]
                                            if wd > mw:
                                                nl = int(len(txt) * (mw / wd)) - 1
                                                self.itemconfig(t,text=txt[:nl])
                                                wd = self.bbox(t)
                                                while wd[2] - wd[0] > mw:
                                                    nl -= 1
                                                    self.dchars(t,nl)
                                                    wd = self.bbox(t)
                                            y += self.xtra_lines_increment
                                            if y + self.half_txt_h > sr:
                                                break
                            except:
                                continue
                elif self.align == "center":
                    for c in range(start_col, end_col - 1):
                        fc = self.col_positions[c]
                        stop = fc + 5
                        sc = self.col_positions[c + 1]
                        mw = sc - fc - 5
                        x = fc + floor(mw / 2)
                        for r in rows_:
                            fr = self.row_positions[r]
                            sr = self.row_positions[r + 1]
                            if sr > sb:
                                sr = sb
                            if (r, self.displayed_columns[c]) in self.highlighted_cells and ((r, c) in self.selected_cells or r in self.selected_rows or c in self.selected_cols):
                                c_1 = self.highlighted_cells[(r, self.displayed_columns[c])][0] if self.highlighted_cells[(r, self.displayed_columns[c])][0].startswith("#") else Color_Map_[self.highlighted_cells[(r, self.displayed_columns[c])][0]]
                                cr_(fc + 1,
                                    fr + 1,
                                    sc,
                                    sr,
                                    fill = (f"#{int((int(c_1[1:3], 16) + c_2_[0]) / 2):02X}" +
                                            f"{int((int(c_1[3:5], 16) + c_2_[1]) / 2):02X}" +
                                            f"{int((int(c_1[5:], 16) + c_2_[2]) / 2):02X}"),
                                    outline = "")
                                tf = self.selected_cells_foreground if self.highlighted_cells[(r, self.displayed_columns[c])][1] is None else self.highlighted_cells[(r, self.displayed_columns[c])][1]
                            elif (r, c) in self.selected_cells:
                                cr_(fc + 1, fr + 1, sc, sr, fill = self.selected_cells_background, outline = "")
                                tf = self.selected_cells_foreground
                            elif (r, self.displayed_columns[c]) in self.highlighted_cells and r not in self.selected_rows and c not in self.selected_cols:
                                cr_(fc + 1, fr + 1, sc, sr, fill = self.highlighted_cells[(r, self.displayed_columns[c])][0], outline = "")
                                tf = self.text_color if self.highlighted_cells[(r, self.displayed_columns[c])][1] is None else self.highlighted_cells[(r, self.displayed_columns[c])][1]
                            elif r in self.selected_rows or c in self.selected_cols:
                                tf = self.selected_cells_foreground
                            else:
                                tf = self.text_color
                            if stop > x2:
                                continue
                            try:
                                lns = self.data_ref[r][self.displayed_columns[c]]
                                if isinstance(lns, str):
                                    lns = lns.split("\n")
                                else:
                                    lns = (f"{lns}", )
                                fl = lns[0]
                                y = fr + self.fl_ins
                                if y + self.half_txt_h > y1:
                                    t = ct_(x, y, text = fl, fill = tf, font = self.my_font, anchor = "center")
                                    wd = self.bbox(t)
                                    wd = wd[2] - wd[0]
                                    if wd > mw:
                                        tl = len(fl)
                                        slce = tl - floor(tl * (mw / wd))
                                        if slce % 2:
                                            slce += 1
                                        else:
                                            slce += 2
                                        slce = int(slce / 2)
                                        fl = fl[slce:tl - slce]
                                        self.itemconfig(t, text = fl)
                                        wd = self.bbox(t)
                                        while wd[2] - wd[0] > mw:
                                            fl = fl[1:-1]
                                            self.itemconfig(t, text = fl)
                                            wd = self.bbox(t)
                                if len(lns) > 1:
                                    stl = int((y1 - y) / self.xtra_lines_increment) - 1
                                    if stl < 1:
                                        stl = 1
                                    y += (stl * self.xtra_lines_increment)
                                    if y + self.half_txt_h < sr:
                                        for i in range(stl, len(lns)):
                                            txt = lns[i]
                                            t = ct_(x, y, text = txt, fill = tf, font = self.my_font, anchor = "center")
                                            wd = self.bbox(t)
                                            wd = wd[2] - wd[0]
                                            if wd > mw:
                                                tl = len(txt)
                                                slce = tl - floor(tl * (mw / wd))
                                                if slce % 2:
                                                    slce += 1
                                                else:
                                                    slce += 2
                                                slce = int(slce / 2)
                                                txt = txt[slce:tl - slce]
                                                self.itemconfig(t, text = txt)
                                                wd = self.bbox(t)
                                                while wd[2] - wd[0] > mw:
                                                    txt = txt[1:-1]
                                                    self.itemconfig(t, text = txt)
                                                    wd = self.bbox(t)
                                            y += self.xtra_lines_increment
                                            if y + self.half_txt_h > sr:
                                                break
                            except:
                                continue
        except:
            return
        if redraw_header:
            self.CH.redraw_grid_and_text(last_col_line_pos, x1, x_stop, start_col, end_col)
        if redraw_row_index:
            self.RI.redraw_grid_and_text(last_row_line_pos, y1, y_stop, start_row, end_row + 1, y2, x1, x_stop)

    def GetColCoords(self, c, sel = False):
        last_col_line_pos = self.col_positions[-1] + 1
        last_row_line_pos = self.row_positions[-1] + 1
        x1 = self.col_positions[c]
        x2 = self.col_positions[c + 1]
        y1 = self.canvasy(0)
        y2 = self.canvasy(self.winfo_height())
        if last_row_line_pos < y2:
            y2 = last_col_line_pos
        if sel:
            return x1, y1 + 1, x2, y2
        else:
            return x1, y1, x2, y2

    def GetRowCoords(self, r, sel = False):
        last_col_line_pos = self.col_positions[-1] + 1
        x1 = self.canvasx(0)
        x2 = self.canvasx(self.winfo_width())
        if last_col_line_pos < x2:
            x2 = last_col_line_pos
        y1 = self.row_positions[r]
        y2 = self.row_positions[r + 1]
        if sel:
            return x1, y1 + 1, x2, y2
        else:
            return x1, y1, x2, y2

    def get_selected_rows(self, get_cells = False):
        if get_cells:
            return sorted(set([cll[0] for cll in self.selected_cells] + list(self.selected_rows)))
        return sorted(self.selected_rows)

    def get_selected_cols(self, get_cells = False):
        if get_cells:
            return sorted(set([cll[0] for cll in self.selected_cells] + list(self.selected_cols)))
        return sorted(self.selected_cols)

    def anything_selected(self, exclude_columns = False, exclude_rows = False, exclude_cells = False):
        if exclude_columns and exclude_rows and not exclude_cells:
            if self.selected_cells:
                return True
        elif exclude_columns and exclude_cells and not exclude_rows:
            if self.selected_rows:
                return True
        elif exclude_rows and exclude_cells and not exclude_columns:
            if self.selected_cols:
                return True
            
        elif exclude_columns and not exclude_rows and not exclude_cells:
            if self.selected_rows or self.selected_cells:
                return True
        elif exclude_rows and not exclude_columns and not exclude_cells:
            if self.selected_cols or self.selected_cells:
                return True
        elif exclude_cells and not exclude_columns and not exclude_rows:
            if self.selected_cols or self.selected_rows:
                return True
            
        elif not exclude_columns and not exclude_rows and not exclude_cells:
            if self.selected_cols or self.selected_rows or self.selected_cells:
                return True
        return False

    def get_selected_cells(self, get_rows = False, get_cols = False):
        res = []
        # IMPROVE SPEED
        if get_cols:
            if self.selected_cols:
                for c in self.selected_cols:
                    for r in range(self.total_rows):
                        res.append((r, c))
        if get_rows:
            if self.selected_rows:
                for r in self.selected_rows:
                    for c in range(self.total_cols):
                        res.append((r, c))
        return res + list(self.selected_cells)

    def edit_cell_(self, event = None):
        if not self.anything_selected():
            return
        if not self.selected_cells:
            return
        st = set("qazwsxedcrfvtgbyhnujmikolpQAZWSXEDCRFVTGBYHNUJMIKOLP0987654321")
        x1 = self.get_min_selected_cell_x()
        y1 = self.get_min_selected_cell_y()
        if event.char in st:
            text = event.char
        else:
            text = self.data_ref[y1][x1]
        self.select_cell(r = y1, c = x1)
        self.see(r = y1, c = x1, keep_yscroll = False, keep_xscroll = False, bottom_right_corner = False, check_cell_visibility = True)
        self.RI.set_row_height(y1, only_set_if_too_small = True)
        self.CH.set_col_width(x1, only_set_if_too_small = True)
        self.refresh()
        self.create_text_editor(r = y1, c = x1, text = text, set_data_ref_on_destroy = True)   
        
    def create_text_editor(self, r = 0, c = 0, text = None, state = "normal", see = True, set_data_ref_on_destroy = False):
        if see:
            self.see(r = r, c = c, check_cell_visibility = True)
        x = self.col_positions[c]
        y = self.row_positions[r]
        w = self.col_positions[c + 1] - x
        h = self.row_positions[r + 1] - y
        if text is None:
            text = ""
        self.text_editor = TextEditor(self, text = text, font = self.my_font, state = state, width = w, height = h)
        self.text_editor_id = self.create_window((x, y), window = self.text_editor, anchor = "nw")
        self.text_editor.textedit.bind("<Alt-Return>", self.text_editor_newline_binding)
        if set_data_ref_on_destroy:
            self.text_editor.textedit.bind("<Return>", lambda x: self.get_text_editor_value((r, c)))
            self.text_editor.textedit.bind("<FocusOut>", lambda x: self.get_text_editor_value((r, c)))
            self.text_editor.textedit.focus_set()

    def bind_text_editor_destroy(self, binding, r, c):
        self.text_editor.textedit.bind("<Return>", lambda x: binding((r, c)))
        self.text_editor.textedit.bind("<FocusOut>", lambda x: binding((r, c)))
        self.text_editor.textedit.focus_set()

    def destroy_text_editor(self):
        try:
            self.delete(self.text_editor_id)
            self.text_editor.destroy()
            self.text_editor_id = None
        except:
            pass
        self.text_editor = None

    def text_editor_newline_binding(self, event = None):
        self.text_editor.config(height = self.text_editor.winfo_height() + self.xtra_lines_increment)

    def get_text_editor_value(self, destroy_tup = None, r = None, c = None, set_data_ref_on_destroy = True, event = None, destroy = True, move_down = True):
        if self.text_editor is not None:
            self.text_editor_value = self.text_editor.get()
        if destroy:
            try:
                self.delete(self.text_editor_id)
                self.text_editor.destroy()
                self.text_editor_id = None
            except:
                pass
            self.text_editor = None
        if set_data_ref_on_destroy:
            if r is None and c is None and destroy_tup:
                r, c = destroy_tup[0], destroy_tup[1]
            self.undo_storage.append(zlib.compress(pickle.dumps({(r, c): f"{self.data_ref[r][c]}"})))
            self.data_ref[r][c] = self.text_editor_value
            self.RI.set_row_height(r)
            self.CH.set_col_width(c, only_set_if_too_small = True)
            if self.extra_edit_cell_func is not None:
                self.extra_edit_cell_func((r, c))
        if move_down:
            self.arrowkey_DOWN()
        self.refresh()
        self.focus_set()
        return self.text_editor_value

    def create_dropdown(self, r = 0, c = 0, values = [], set_value = None, state = "readonly", see = True, destroy_on_select = True, current = False):
        if see:
            if not self.cell_is_completely_visible(r = r, c = c):
                self.see(r = r, c = c, check_cell_visibility = False)
        x = self.col_positions[c]
        y = self.row_positions[r]
        w = self.GetWidthChars(self.col_positions[c + 1] - x)
        self.table_dropdown = TableDropdown(self, font = self.my_font, state = state, values = values, set_value = set_value, width = w)
        self.table_dropdown_id = self.create_window((x, y), window = self.table_dropdown, anchor = "nw")
        if destroy_on_select:
            self.table_dropdown.bind("<<ComboboxSelected>>", lambda event: self.get_dropdown_value(current = current))

    def get_dropdown_value(self, event = None, current = False, destroy = True):
        if self.table_dropdown is not None:
            if current:
                self.table_dropdown_value = self.table_dropdown.current()
            else:
                self.table_dropdown_value = self.table_dropdown.get_my_value()
        if destroy:
            try:
                self.delete(self.table_dropdown_id)
                self.table_dropdown.destroy()
            except:
                pass
            self.table_dropdown = None
        return self.table_dropdown_value
    

class TableDropdown(ttk.Combobox):
    def __init__(self, parent, font, state, values = [], set_value = None, width = None):
        self.displayed = tk.StringVar()
        ttk.Combobox.__init__(self,
                              parent,
                              font = font,
                              state = state,
                              values = values,
                              set_value = set_value,
                              textvariable = self.displayed)
        if width:
            self.config(width = width)
        if set_value is not None:
            self.displayed.set(set_value)
        elif values:
            self.displayed.set(values[0])
            
    def get_my_value(self, event = None):
        return self.displayed.get()
    
    def set_my_value(self, value, event = None):
        self.displayed.set(value)
        
        
class Sheet(tk.Frame):
    def __init__(self,
                 C,
                 show = True,
                 width = None,
                 height = None,
                 headers = None,
                 header_background = "white",
                 header_border_color = "#ababab",
                 header_grid_color = "#ababab",
                 header_foreground = "black",
                 header_select_background = "#707070",
                 header_select_foreground = "white",
                 data_reference = None,
                 column_width = 200,
                 header_height = "1",
                 max_colwidth = "inf",
                 max_rh = "inf",
                 max_header_height = "inf",
                 max_row_width = "inf",
                 row_index = None,
                 row_index_width = 100,
                 row_height = "1",
                 row_index_background = "white",
                 row_index_border_color = "#ababab",
                 row_index_grid_color = "#ababab",
                 row_index_foreground = "black",
                 row_index_select_background = "#707070",
                 row_index_select_foreground = "white",
                 top_left_background = "white",
                 top_left_foreground = "gray85",
                 font = ("TkDefaultFont", 10, "normal"),
                 header_font = ("TkHeadingFont", 11, "bold"),
                 align = "w",
                 header_align = "center",
                 row_index_align = "center",
                 table_background = "white",
                 grid_color = "#d4d4d4",
                 text_color = "black",
                 selected_cells_background = "#707070",
                 selected_cells_foreground = "white",
                 resizing_line_color = "black",
                 drag_and_drop_color = "turquoise1",
                 displayed_columns = [],
                 all_columns_displayed = True,
                 outline_thickness = 0,
                 outline_color = "gray2",
                 theme = "normal"):
        tk.Frame.__init__(self,
                          C,
                          width = width,
                          height = height,
                          highlightthickness = outline_thickness,
                          highlightbackground = outline_color)
        self.C = C
        self.grid_columnconfigure(1, weight = 1)
        self.grid_rowconfigure(1, weight = 1)
        self.RI = RowIndexes(self,
                             max_rh = max_rh,
                             max_row_width = max_row_width,
                             row_index_width = row_index_width,
                             row_index_align = row_index_align,
                             row_index_background = row_index_background,
                             row_index_border_color = row_index_border_color,
                             row_index_grid_color = row_index_grid_color,
                             row_index_foreground = row_index_foreground,
                             row_index_select_background = row_index_select_background,
                             row_index_select_foreground = row_index_select_foreground,
                             drag_and_drop_color = drag_and_drop_color,
                             resizing_line_color = resizing_line_color)
        self.CH = ColumnHeaders(self,
                                max_colwidth = max_colwidth,
                                max_header_height = max_header_height,
                                header_align = header_align,
                                header_background = header_background,
                                header_border_color = header_border_color,
                                header_grid_color = header_grid_color,
                                header_foreground = header_foreground,
                                header_select_background = header_select_background,
                                header_select_foreground = header_select_foreground,
                                drag_and_drop_color = drag_and_drop_color,
                                resizing_line_color = resizing_line_color)
        self.MT = MainTable(self,
                            column_width = column_width,
                            row_height = row_height,
                            column_headers_canvas = self.CH,
                            row_index_canvas = self.RI,
                            headers = headers,
                            header_height = header_height,
                            data_reference = data_reference,
                            row_index = row_index,
                            font = font,
                            header_font = header_font,
                            align = align,
                            table_background = table_background,
                            grid_color = grid_color,
                            text_color= text_color,
                            selected_cells_background = selected_cells_background,
                            selected_cells_foreground = selected_cells_foreground,
                            displayed_columns = displayed_columns,
                            all_columns_displayed = all_columns_displayed)
        self.TL = TopLeftRectangle(parentframe = self,
                                   main_canvas = self.MT,
                                   row_index_canvas = self.RI,
                                   header_canvas = self.CH,
                                   background = top_left_background,
                                   foreground = top_left_foreground)
        if theme != "normal":
            self.change_theme(theme)
        self.yscroll = ttk.Scrollbar(self, command = self.MT.set_yviews, orient = "vertical")
        self.xscroll = ttk.Scrollbar(self, command = self.MT.set_xviews, orient = "horizontal")
        self.MT["xscrollcommand"] = self.xscroll.set
        self.MT["yscrollcommand"] = self.yscroll.set
        self.CH["xscrollcommand"] = self.xscroll.set
        self.RI["yscrollcommand"] = self.yscroll.set
        if show:
            self.TL.grid(row = 0, column = 0)
            self.RI.grid(row = 1, column = 0, sticky = "nswe")
            self.CH.grid(row = 0, column = 1, sticky = "nswe")
            self.MT.grid(row = 1, column = 1, sticky = "nswe")
            self.yscroll.grid(row = 1, column = 2, sticky = "nswe")
            self.xscroll.grid(row = 2, column = 1, columnspan = 2, sticky = "nswe")
            self.MT.update()

    def focus_set(self, canvas = "table"):
        if canvas == "table":
            self.MT.focus_set()
        elif canvas == "header":
            self.CH.focus_set()
        elif canvas == "index":
            self.RI.focus_set()
        elif canvas == "topleft":
            self.TL.focus_set()

    def extra_bindings(self, bindings):
        for binding, func in bindings:
            if binding == "ctrl_c":
                self.MT.extra_ctrl_c_func = func
            if binding == "ctrl_x":
                self.MT.extra_ctrl_x_func = func
            if binding == "ctrl_v":
                self.MT.extra_ctrl_v_func = func
            if binding == "ctrl_z":
                self.MT.extra_ctrl_z_func = func
            if binding == "delete_key":
                self.MT.extra_delete_key_func = func
            if binding == "edit_cell":
                self.MT.extra_edit_cell_func = func
            if binding == "row_index_drag_drop":
                self.RI.ri_extra_drag_drop_func = func
            if binding == "column_header_drag_drop":
                self.CH.ch_extra_drag_drop_func = func
            if binding == "cell_select":
                self.MT.selection_binding_func = func
            if binding == "drag_select":
                self.MT.drag_selection_binding_func = func
                self.RI.drag_selection_binding_func = func
                self.CH.drag_selection_binding_func = func
            if binding == "row_select":
                self.RI.selection_binding_func = func
            if binding == "column_select":
                self.CH.selection_binding_func = func

    def bind(self, binding, func):
        if binding == "<ButtonPress-1>":
            self.MT.extra_b1_press_func = func
            self.CH.extra_b1_press_func = func
            self.RI.extra_b1_press_func = func
            self.TL.extra_b1_press_func = func
        elif binding == "<ButtonMotion-1>":
            self.MT.extra_b1_motion_func = func
            self.CH.extra_b1_motion_func = func
            self.RI.extra_b1_motion_func = func
            self.TL.extra_b1_motion_func = func
        elif binding == "<ButtonRelease-1>":
            self.MT.extra_b1_release_func = func
            self.CH.extra_b1_release_func = func
            self.RI.extra_b1_release_func = func
            self.TL.extra_b1_release_func = func
        elif binding == "<Double-Button-1>":
            self.MT.extra_double_b1_func = func
            self.CH.extra_double_b1_func = func
            self.RI.extra_double_b1_func = func
            self.TL.extra_double_b1_func = func
        elif binding == "<Motion>":
            self.MT.extra_motion_func = func
            self.CH.extra_motion_func = func
            self.RI.extra_motion_func = func
            self.TL.extra_motion_func = func
        else:
            self.MT.bind(binding, func)
            self.CH.bind(binding, func)
            self.RI.bind(binding, func)
            self.TL.bind(binding, func)

    def unbind(self, binding):
        if binding == "<ButtonPress-1>":
            self.MT.extra_b1_press_func = None
            self.CH.extra_b1_press_func = None
            self.RI.extra_b1_press_func = None
            self.TL.extra_b1_press_func = None
        elif binding == "<ButtonMotion-1>":
            self.MT.extra_b1_motion_func = None
            self.CH.extra_b1_motion_func = None
            self.RI.extra_b1_motion_func = None
            self.TL.extra_b1_motion_func = None
        elif binding == "<ButtonRelease-1>":
            self.MT.extra_b1_release_func = None
            self.CH.extra_b1_release_func = None
            self.RI.extra_b1_release_func = None
            self.TL.extra_b1_release_func = None
        elif binding == "<Double-Button-1>":
            self.MT.extra_double_b1_func = None
            self.CH.extra_double_b1_func = None
            self.RI.extra_double_b1_func = None
            self.TL.extra_double_b1_func = None
        elif binding == "<Motion>":
            self.MT.extra_motion_func = None
            self.CH.extra_motion_func = None
            self.RI.extra_motion_func = None
            self.TL.extra_motion_func = None
        else:
            self.MT.unbind(binding)
            self.CH.unbind(binding)
            self.RI.unbind(binding)
            self.TL.unbind(binding)

    def basic_bindings(self, enable = False):
        if enable:
            for canvas in (self.MT, self.CH, self.RI, self.TL):
                canvas.basic_bindings(onoff = "enable")
        elif not enable:
            for canvas in (self.MT, self.CH, self.RI, self.TL):
                canvas.basic_bindings(onoff = "disable")

    def edit_bindings(self, enable = False):
        if enable:
            self.MT.edit_bindings(onoff = "enable")
        elif not enable:
            self.MT.edit_bindings(onoff = "disable")

    def cell_edit_binding(self, enable = False):
        self.MT.bind_cell_edit(enable)

    def identify_region(self, event):
        # UNFINISHED, ADD SEPARATOR?
        if event.widget == self.MT:
            return "table"
        elif event.widget == self.RI:
            return "index"
        elif event.widget == self.CH:
            return "header"
        elif event.widget == self.TL:
            return "top left"

    def identify_row(self, event, exclude_index = False, allow_end = True):
        ev_w = event.widget
        if ev_w == self.MT:
            return self.MT.identify_row(y = event.y, allow_end = allow_end)
        elif ev_w == self.RI:
            if exclude_index:
                return None
            else:
                return self.MT.identify_row(y = event.y, allow_end = allow_end)
        elif ev_w == self.CH or ev_w == self.TL:
            return None

    def identify_column(self, event, exclude_header = False, allow_end = True):
        ev_w = event.widget
        if ev_w == self.MT:
            return self.MT.identify_col(x = event.x, allow_end = allow_end)
        elif ev_w == self.RI or ev_w == self.TL:
            return None
        elif ev_w == self.CH:
            if exclude_header:
                return None
            else:
                return self.MT.identify_col(x = event.x, allow_end = allow_end)

    def get_example_canvas_column_widths(self, total_cols = None):
        colpos = int(self.MT.default_cw)
        if total_cols is not None:
            return [0] + list(accumulate(colpos for c in range(total_cols)))
        return [0] + list(accumulate(colpos for c in range(self.MT.total_cols)))

    def get_example_canvas_row_heights(self, total_rows = None):
        rowpos = self.MT.GetLinesHeight(self.MT.default_rh)
        if total_rows is not None:
            return [0] + list(accumulate(rowpos for c in range(total_rows)))
        return [0] + list(accumulate(rowpos for c in range(self.MT.total_rows)))

    def get_column_widths(self, canvas_positions = False):
        if canvas_positions:
            return [int(n) for n in self.MT.col_positions]
        return [int(b - a) for a,b in zip(self.MT.col_positions, islice(self.MT.col_positions, 1, len(self.MT.col_positions)))]
    
    def get_row_heights(self, canvas_positions = False):
        if canvas_positions:
            return [int(n) for n in self.MT.row_positions]
        return [int(b - a) for a, b in zip(self.MT.row_positions, islice(self.MT.row_positions, 1, len(self.MT.row_positions)))]

    def column_width(self, column = None, width = None, only_set_if_too_small = False, redraw = True):
        if column == "all":
            if width == "default":
                self.MT.reset_col_positions()
        elif column == "displayed":
            if width == "text":
                sc, ec = self.MT.get_visible_columns(self.MT.canvasx(0), self.MT.canvasx(self.winfo_width()))
                for c in range(sc, ec - 1):
                    self.CH.set_col_width(c)
        elif width is not None and column is not None:
            self.CH.set_col_width(col = column, width = width, only_set_if_too_small = only_set_if_too_small)
        elif column is not None:
            return int(self.MT.col_positions[column + 1] - self.MT.col_positions[column])
        if redraw:
            self.refresh()

    def set_column_widths(self, column_widths = None, canvas_positions = False, reset = False, verify = False):
        cwx = None
        if reset:
            self.MT.reset_col_positions()
            return
        if verify:
            cwx = self.verify_column_widths(column_widths, canvas_positions)
        if isinstance(column_widths, list):
            if canvas_positions:
                self.MT.col_positions = column_widths
            else:
                self.MT.col_positions = [0] + list(accumulate(width for width in column_widths))
        return cwx
            
    def set_row_height(self, row = None, height = None):
        if row == "all":
            if height == "default":
                self.MT.reset_row_positions()
        else:
            self.MT.row_positions[row + 1] = height

    def set_row_heights(self, row_heights = None, canvas_positions = False, reset = False, verify = False):
        rhx = None
        if reset:
            self.MT.reset_row_positions()
            return
        if verify:
            rhx = self.verify_row_heights(row_heights, canvas_positions)
        if isinstance(row_heights, list):
            if canvas_positions:
                self.MT.row_positions = row_heights
            else:
                self.MT.row_positions = [0] + list(accumulate(height for height in row_heights))
        return rhx

    def verify_row_heights(self, row_heights, canvas_positions = False):
        if row_heights[0] != 0 or isinstance(row_heights[0],bool):
            return False
        if not isinstance(row_heights,list):
            return False
        if canvas_positions:
            if any(x - z < self.MT.min_rh or not isinstance(x, int) or isinstance(x, bool) for z, x in zip(islice(row_heights, 0, None), islice(row_heights, 1, None))):
                return False
        elif not canvas_positions:
            if any(z < self.MT.min_rh or not isinstance(z, int) or isinstance(z, bool) for z in row_heights):
                return False
        return True

    def verify_column_widths(self, column_widths, canvas_positions = False):
        if column_widths[0] != 0 or isinstance(column_widths[0],bool):
            return False
        if not isinstance(column_widths,list):
            return False
        if canvas_positions:
            if any(x - z < self.MT.min_cw or not isinstance(x, int) or isinstance(x, bool) for z, x in zip(islice(column_widths, 0, None), islice(column_widths, 1, None))):
                return False
        elif not canvas_positions:
            if any(z < self.MT.min_cw or not isinstance(z, int) or isinstance(z, bool) for z in column_widths):
                return False
        return True
                
    def default_row_height(self, height = None):
        if height is not None:
            self.MT.default_rh = int(height)
        return self.MT.default_rh

    def default_header_height(self, height = None):
        if height is not None:
            self.MT.default_hh = int(height)
        return self.MT.default_hh

    def create_dropdown(self,
                        r = 0,
                        c = 0,
                        values = [],
                        set_value = None,
                        state = "readonly",
                        see = True,
                        destroy_on_select = True,
                        current = False):
        self.MT.create_dropdown(r = r, c = c, values = values, set_value = set_value, state = state, see = see,
                                destroy_on_select = destroy_on_select, current = current)

    def get_dropdown_value(self, current = False, destroy = True):
        return self.MT.get_dropdown_value(current = current, destroy = destroy)

    def delete_row_position(self, idx, deselect_all = False, preserve_other_selections = False):
        self.MT.del_row_position(idx = idx,
                                 deselect_all = deselect_all,
                                 preserve_other_selections = preserve_other_selections)

    def delete_row(self, idx = -1, deselect_all = False, preserve_other_selections = False):
        del self.MT.data_ref[idx]
        self.MT.del_row_position(idx = idx,
                                 deselect_all = deselect_all,
                                 preserve_other_selections = preserve_other_selections)

    def insert_row_position(self, idx, height = None, deselect_all = False, preserve_other_selections = False):
        self.MT.insert_row_position(idx = idx, height = height,
                                    deselect_all = deselect_all,
                                    preserve_other_selections = preserve_other_selections)

    def insert_row(self, row = None, idx = "end", height = None, deselect_all = False, preserve_other_selections = False):
        self.MT.insert_row_position(idx = idx, height = height,
                                    deselect_all = deselect_all,
                                    preserve_other_selections = preserve_other_selections)
        if idx.lower() == "end":
            if isinstance(row, list):
                self.MT.data_ref.append(row if row else list(repeat("", len(self.MT.col_positions) - 1)))
            else:
                self.MT.data_ref.append([e for e in row])
        else:
            if isinstance(row, list):
                self.MT.data_ref.insert(idx, row if row else list(repeat("", len(self.MT.col_positions) - 1)))
            else:
                self.MT.data_ref.insert(idx, [e for e in row])

    def move_row_position(self, row, moveto):
        self.MT.move_row_position(row, moveto)

    def move_row(self, row, moveto):
        self.MT.move_row_position(row, moveto)
        self.MT.data_ref.insert(moveto, self.MT.data_ref.pop(row))

    def delete_column_position(self, idx, deselect_all = False, preserve_other_selections = False):
        self.MT.del_col_position(idx,
                                 deselect_all = deselect_all,
                                 preserve_other_selections = preserve_other_selections)

    def delete_column(self, idx = -1, deselect_all = False, preserve_other_selections = False):
        for rn in range(len(self.MT.data_ref)):
            del self.MT.data_ref[rn][idx] 
        self.MT.del_col_position(idx,
                                 deselect_all = deselect_all,
                                 preserve_other_selections = preserve_other_selections)

    def insert_column_position(self, idx, width = None, deselect_all = False, preserve_other_selections = False):
        self.MT.insert_col_position(idx = idx,
                                    width = width,
                                    deselect_all = deselect_all,
                                    preserve_other_selections = preserve_other_selections)

    def insert_column(self, column = None, idx = "end", width = None, deselect_all = False, preserve_other_selections = False):
        self.MT.insert_col_position(idx = idx,
                                    width = width,
                                    deselect_all = deselect_all,
                                    preserve_other_selections = preserve_other_selections)
        if idx.lower() == "end":
            for rn, col_value in zip(range(len(self.MT.data_ref)), column):
                self.MT.data_ref[rn].append(col_value)
        else:
            for rn, col_value in zip(range(len(self.MT.data_ref)), column):
                self.MT.data_ref[rn].insert(idx, col_value)

    def move_column_position(self, column, moveto):
        self.MT.move_col_position(column, moveto)

    def move_column(self, column, moveto):
        self.MT.move_col_position(column, moveto)
        for rn in range(len(self.MT.data_ref)):
            self.MT.data_ref[rn].insert(moveto, self.MT.data_ref[rn].pop(column))

    def create_text_editor(self, row = 0, column = 0, text = None, state = "normal", see = True, set_data_ref_on_destroy = False):
        self.MT.create_text_editor(r = row, c = column, text = text, state = state, see = see, set_data_ref_on_destroy = set_data_ref_on_destroy)

    def get_xview(self):
        return self.MT.xview()

    def get_yview(self):
        return self.MT.yview()

    def set_xview(self, position, option = "moveto"):
        self.MT.set_xviews(option, position)

    def set_yview(self,position,option = "moveto"):
        self.MT.set_yviews(option, position)

    def set_view(self, x_args, y_args):
        self.MT.set_view(x_args, y_args)

    def see(self, row = 0, column = 0, keep_yscroll = False, keep_xscroll = False, bottom_right_corner = False, check_cell_visibility = True):
        self.MT.see(row, column, keep_yscroll, keep_xscroll, bottom_right_corner, check_cell_visibility = check_cell_visibility)

    def select_row(self, row, redraw = True):
        self.RI.select_row(row, redraw = redraw)

    def select_column(self, column, redraw = True):
        self.CH.select_col(column, redraw = redraw)

    def select_cell(self, row, column, redraw = True):
        self.MT.select_cell(row, column, redraw = redraw)

    def add_cell_selection(self, r, c, redraw = True, run_binding_func = True):
        self.MT.add_selection(r = r, c = c, redraw = redraw, run_binding_func = run_binding_func)

    def add_row_selection(self, r, redraw = True, run_binding_func = True):
        self.RI.add_selection(r = r, redraw = redraw, run_binding_func = run_binding_func)

    def add_column_selection(self, c, redraw = True, run_binding_func = True):
        self.CH.add_selection(c = c, redraw = redraw, run_binding_func = run_binding_func)

    def select(self, row = None, column = None, cell = None, redraw = True):
        self.MT.select(r = row, c = column, cell = cell, redraw = redraw)

    def deselect(self, row = None, column = None, cell = None, redraw = True):
        self.MT.deselect(r = row, c = column, cell = cell, redraw = redraw)

    def get_selected_rows(self, get_cells = False):
        return self.MT.get_selected_rows(get_cells = get_cells)

    def get_selected_columns(self, get_cells = False):
        return self.MT.get_selected_cols(get_cells = get_cells)

    def get_selected_cells(self, get_rows = False, get_cols = False):
        return self.MT.get_selected_cells(get_rows = get_rows, get_cols = get_cols)

    def anything_selected(self, exclude_columns = False, exclude_rows = False, exclude_cells = False):
        return self.MT.anything_selected(exclude_columns = exclude_columns, exclude_rows = exclude_rows, exclude_cells = exclude_cells)

    def highlight_cells(self, row = 0, column = 0, cells = [], canvas = "table", bg = None, fg = None, redraw = False):
        if canvas == "table":
            self.MT.highlight_cells(r = row,
                                    c = column,
                                    cells = cells,
                                    bg = bg,
                                    fg = fg,
                                    redraw = redraw)
        elif canvas == "row_index":
            self.RI.highlight_cells(r = row,
                                    cells = cells,
                                    bg = bg,
                                    fg = fg,
                                    redraw = redraw)
        elif canvas == "header":
            self.CH.highlight_cells(c = column,
                                    cells = cells,
                                    bg = bg,
                                    fg = fg,
                                    redraw = redraw)

    def dehighlight_cells(self, row = 0, column = 0, cells = [], canvas = "table", all_ = False, redraw = True):
        if canvas == "table":
            if cells and not all_:
                for t in cells:
                    try:
                        del self.MT.highlighted_cells[t]
                    except:
                        pass
            elif not all_:
                if (row, column) in self.MT.highlighted_cells:
                    del self.MT.highlighted_cells[(row, column)]
            elif all_:
                self.MT.highlighted_cells = {}
        elif canvas == "row_index":
            if cells and not all_:
                for r in cells:
                    try:
                        del self.RI.highlighted_cells[r]
                    except:
                        pass
            elif not all_:
                if row in self.RI.highlighted_cells:
                    del self.RI.highlighted_cells[row]
            elif all_:
                self.RI.highlighted_cells = {}
        elif canvas == "header":
            if cells and not all_:
                for c in cells:
                    try:
                        del self.CH.highlighted_cells[c]
                    except:
                        pass
            elif not all_:
                if column in self.CH.highlighted_cells:
                    del self.CH.highlighted_cells[column]
            elif all_:
                self.CH.highlighted_cells = {}
        if redraw:
            self.refresh(True, True)

    def get_highlighted_cells(self, canvas = "table"):
        if canvas == "table":
            return dict(tuple(tup) for tup in self.MT.highlighted_cells)
        
    def get_frame_y(self, y):
        return y + self.CH.current_height

    def get_frame_x(self, x):
        return x + self.RI.current_width

    def align(self, align = None, redraw = True):
        if align is None:
            return self.MT.align
        elif align in ("w", "center"):
            self.MT.align = align
        else:
            raise ValueError("Align must be either 'w' or 'center'")
        if redraw:
            self.refresh()

    def header_align(self, align = None, redraw = True):
        if align is None:
            return self.CH.align
        elif align in ("w", "center"):
            self.CH.align = align
        else:
            raise ValueError("Align must be either 'w' or 'center'")
        if redraw:
            self.refresh()

    def row_index_align(self, align = None, redraw = True):
        if align is None:
            return self.RI.align
        elif align in ("w", "center"):
            self.RI.align = align
        else:
            raise ValueError("Align must be either 'w' or 'center'")
        if redraw:
            self.refresh()

    def font(self,newfont=None):
        self.MT.font(newfont)

    def header_font(self, newfont = None):
        self.MT.header_font(newfont)

    def change_color(self,
                     header_background = None,
                     header_border_color = None,
                     header_grid_color = None,
                     header_foreground = None,
                     header_select_background = None,
                     header_select_foreground = None,
                     row_index_background = None,
                     row_index_border_color = None,
                     row_index_grid_color = None,
                     row_index_foreground = None,
                     row_index_select_background = None,
                     row_index_select_foreground = None,
                     top_left_background = None,
                     top_left_foreground = None,
                     table_background = None,
                     grid_color = None,
                     text_color = None,
                     selected_cells_background = None,
                     selected_cells_foreground = None,
                     resizing_line_color = None,
                     drag_and_drop_color = None,
                     outline_thickness = None,
                     outline_color = None,
                     redraw = False):
        if header_background is not None:
            self.CH.config(background = header_background)
        if header_border_color is not None:
            self.CH.header_border_color = header_border_color
        if header_grid_color is not None:
            self.CH.grid_color = header_grid_color
        if header_foreground is not None:
            self.CH.text_color = header_foreground
        if header_select_background is not None:
            self.CH.selected_cells_background = header_select_background
        if header_select_foreground is not None:
            self.CH.selected_cells_foreground = header_select_foreground
        if row_index_background is not None:
            self.RI.config(background=row_index_background)
        if row_index_border_color is not None:
            self.RI.row_index_border_color = row_index_border_color
        if row_index_grid_color is not None:
            self.RI.grid_color = row_index_grid_color
        if row_index_foreground is not None:
            self.RI.text_color = row_index_foreground
        if row_index_select_background is not None:
            self.RI.selected_cells_background = row_index_select_background
        if row_index_select_foreground is not None:
            self.RI.selected_cells_foreground = row_index_select_foreground
        if top_left_background is not None:
            self.TL.config(background=top_left_background)
        if top_left_foreground is not None:
            self.TL.rectangle_foreground = top_left_foreground
            self.TL.itemconfig("rw", fill = top_left_foreground)
            self.TL.itemconfig("rh", fill = top_left_foreground)
        if table_background is not None:
            self.MT.config(background = table_background)
        if grid_color is not None:
            self.MT.grid_color = grid_color
        if text_color is not None:
            self.MT.text_color = text_color
        if selected_cells_background is not None:
            self.MT.selected_cells_background = selected_cells_background
        if selected_cells_foreground is not None:
            self.MT.selected_cells_foreground = selected_cells_foreground
        if resizing_line_color is not None:
            self.CH.resizing_line_color = resizing_line_color
            self.RI.resizing_line_color = resizing_line_color
        if drag_and_drop_color is not None:
            self.CH.drag_and_drop_color = drag_and_drop_color
            self.RI.drag_and_drop_color = drag_and_drop_color
        if outline_thickness is not None:
            self.config(highlightthickness = outline_thickness)
        if outline_color is not None:
            self.config(highlightbackground = outline_color)
        if redraw:
            self.refresh()

    def change_theme(self, theme = "normal"):
        if theme == "normal":
            self.change_color(header_background = "white",
                                header_border_color = "#ababab",
                                header_grid_color = "#ababab",
                                header_foreground = "black",
                                header_select_background = "#707070",
                                header_select_foreground = "white",
                                row_index_background = "white",
                                row_index_border_color = "#ababab",
                                row_index_grid_color = "#ababab",
                                row_index_foreground = "black",
                                row_index_select_background = "#707070",
                                row_index_select_foreground = "white",
                                top_left_background = "white",
                                top_left_foreground = "gray85",
                                table_background = "white",
                                grid_color = "#d4d4d4",
                                text_color = "black",
                                selected_cells_background = "#707070",
                                selected_cells_foreground = "white",
                                resizing_line_color = "black",
                                drag_and_drop_color = "turquoise1",
                                outline_color = "gray2",
                                redraw = True)
        elif theme == "dark":
            self.change_color(header_background = "#222222",
                                header_border_color = "#353a41",
                                header_grid_color = "#353a41",
                                header_foreground = "gray95",
                                header_select_background = "#7a8288",
                                header_select_foreground = "black",
                                row_index_background = "#222222",
                                row_index_border_color = "#353a41",
                                row_index_grid_color = "#353a41",
                                row_index_foreground = "gray95",
                                row_index_select_background = "#7a8288",
                                row_index_select_foreground = "black",
                                top_left_background = "#353a41",
                                top_left_foreground = "#222222",
                                table_background = "#222222",
                                grid_color = "#353a41",
                                text_color = "gray95",
                                selected_cells_background = "#7a8288",
                                selected_cells_foreground = "black",
                                resizing_line_color = "white",
                                drag_and_drop_color = "#9acd32",
                                outline_color = "gray2",
                                redraw = True)
            
    def data_reference(self,
                       newdataref = None,
                       total_cols = None,
                       total_rows = None,
                       reset_col_positions = True,
                       reset_row_positions = True,
                       redraw = False):
        return self.MT.data_reference(newdataref,
                                      total_cols,
                                      total_rows,
                                      reset_col_positions,
                                      reset_row_positions,
                                      redraw)

    def display_columns(self, indexes = None, enable = None, reset_col_positions = True, number_of_columns = None, refresh = False, deselect_all = True):
        res = self.MT.display_columns(indexes = indexes,
                                      enable = enable,
                                      reset_col_positions = reset_col_positions,
                                      number_of_columns = number_of_columns,
                                      deselect_all = deselect_all)
        if refresh:
            self.refresh()
        return res

    def show_ctrl_outline(self, t = "cut", canvas = "table", start_cell = (0, 0), end_cell = (1, 1)):
        self.MT.show_ctrl_outline(t = t, canvas = canvas, start_cell = start_cell, end_cell = end_cell)

    def get_max_selected_cell_x(self):
        return self.MT.get_max_selected_cell_x()

    def get_max_selected_cell_y(self):
        return self.MT.get_max_selected_cell_y()

    def get_min_selected_cell_y(self):
        return self.MT.get_min_selected_cell_y()

    def get_min_selected_cell_x(self):
        return self.MT.get_min_selected_cell_x()
        
    def headers(self, newheaders = None, index = None):
        return self.MT.headers(newheaders, index)
        
    def row_index(self, newindex = None, index = None):
        return self.MT.row_index(newindex,index)

    def reset_undos(self):
        self.MT.undo_storage = deque(maxlen = 20)

    def enable_bindings(self, bindings):
        self.MT.enable_bindings(bindings)

    def disable_bindings(self, bindings):
        self.MT.disable_bindings(bindings)

    def show(self, canvas = "all"):
        if canvas == "all":
            self.TL.grid(row = 0, column = 0)
            self.RI.grid(row = 1, column = 0, sticky = "nswe")
            self.CH.grid(row = 0, column = 1, sticky = "nswe")
            self.MT.grid(row = 1, column = 1, sticky = "nswe")
            self.yscroll.grid(row = 0, column = 2, rowspan = 3, sticky = "nswe")
            self.xscroll.grid(row = 2, column = 1, sticky = "nswe")
        elif canvas == "row_index":
            self.RI.grid(row = 1, column = 0, sticky = "nswe")
        elif canvas == "header":
            self.CH.grid(row = 0, column = 1, sticky = "nswe")
        self.MT.update()

    def hide(self, canvas = "all"):
        if canvas.lower() == "all":
            self.TL.grid_forget()
            self.RI.grid_forget()
            self.CH.grid_forget()
            self.MT.grid_forget()
            self.yscroll.grid_forget()
            self.xscroll.grid_forget()
        elif canvas == "row_index":
            self.RI.grid_forget()
        elif canvas == "header":
            self.CH.grid_forget()

    def refresh(self, redraw_header = True, redraw_row_index = True):
        self.MT.main_table_redraw_grid_and_text(redraw_header = redraw_header, redraw_row_index = redraw_row_index)


if __name__ == '__main__':
    class demo(tk.Tk):
        def __init__(self):
            tk.Tk.__init__(self)
            self.state("zoomed")
            self.grid_columnconfigure(0, weight = 1)
            self.grid_rowconfigure(0, weight = 1)
            self.data = [[f"Row {r} Column {c}" for c in range(100)] for r in range(5000)]
            self.sdem = Sheet(self,
                              width = 1000,
                              height = 700,
                              align = "w",
                              header_align = "center",
                              row_index_align = "center",
                              column_width = 180,
                              row_index_width = 50,
                              data_reference = self.data,
                              headers=[f"Header {c}" for c in range(100)])
            self.sdem.enable_bindings(("single",
                                       "drag_select",
                                       "column_drag_and_drop",
                                       "row_drag_and_drop",
                                       "column_select",
                                       "row_select",
                                       "column_width_resize",
                                       "double_click_column_resize",
                                       "row_width_resize",
                                       "column_height_resize",
                                       "arrowkeys",
                                       "row_height_resize",
                                       "double_click_row_resize"))
            self.sdem.edit_bindings(True)
            self.sdem.grid(row = 0, column = 0, sticky = "nswe")
            self.sdem.highlight_cells(row = 0, column = 0, bg = "orange", fg = "blue")
            self.sdem.highlight_cells(row = 0, bg = "orange", fg = "blue", canvas = "row_index")
            self.sdem.highlight_cells(column = 0, bg = "orange", fg = "blue", canvas = "header")


    app = demo()
    app.mainloop()



