# -*- encoding:utf-8 -*-
''' readV11xls.py - びびび.xls を読んで衣装リストを整形する
  コマンドラインオプションは readV11xls.py -h で。
  CSV形式、Wiki用テーブル形式で出力可能
'''

# Python2.7標準以外のパッケージでxlrdの追加インストール要
#  >> pip install xlrd

# for pydoc
__author__ = "TPK"
__version__ = "1.5.0"
__date__    = "20160815"


import xlrd
import codecs
import sys

dbg = 0

def isnum(x):
    if (type(x) is float) | (type(x) is int):
        return True
    else:
        return False

def fitos(f):
    if isnum(f):
        if f == int(f):
            r = u'%d'%f
        else:
            r = u'%.1f'%f
    else:
        r = u''

    return r

# Player Basic Data from Player sheet
plPPos = [
    ('Phase',1),  #リリース区分
    ('PName',2),  #プレイヤー名
    ('PKana',3),  #プレイヤー名読み
    ('PRoma',4),  #プレイヤー名英字
    ('IniRare',5),  #初期レア度
    ('ST1',6),    #ST(Lv1)
    ('ST20',7),   #ST(Lv20 = ☆1'Rookie' max)
    ('ST30',8),   #ST(Lv30 = ☆2'Elite' max)
    ('ST40',9),   #ST(Lv40 = ☆3'Genius' max)
    ('ST50',10),  #ST(Lv50 = ☆4'Fantasista' max)
    ('ST60',11),  #ST(Lv60 = ☆5'Legend' max)
    ('ST70',12),  #ST(Lv70 = ☆6'Venus' max)
    ('SH',13),    #Shoot in-file(rare=CosR) to internal(Rare=1)
    ('DR',14),    #Dribble
    ('PA',15),    #Pass
    ('DF',16),    #Defence
    ('SHMax',17), #Max Sh(Rare=6)
    ('DRMax',18),
    ('PAMax',19),
    ('DFMax',20),
    ('Pos',22),   #適性ポジション
    ('MOfic',23), #経営；会社
    ('MStad',24), #経営；スタジアム
    ('MClub',25), #経営；クラブハウス
    ('PHand',26), #Prof：ニックネーム
    ('PHeig',27), #Prof：身長
    ('PWeig',28), #Prof：身長
    ('PBust',29), #Prof：B
    ('PWais',30), #Prof：W
    ('PHip',31),  #Prof：H
    ('PBirt',32), #Prof：誕生日
    ('PBlod',33), #Prof：血液型
    ('PLoca',34), #Prof：出身
    ('PFavo',35), #Prof：趣味
    ('PCV',36),   #Prof：声優
    ('PNote',37), #Note；ノート(BaseCos)
    ('PIntr',56), #Note；紹介文
    ('SVSkl',38), #SV：スキル
    ('SVSkd',39), #SV：SVスキル効果
    ('SVCvt',40), #SV；コンバートPos
    ('SVSta',41), #SV：スタミナ上限(Lv80)
    ('SVSH',42), #SV：SH上限(Lv80)
    ('SVDR',43), #SV：DR上限(Lv80)
    ('SVPA',44), #SV：PA上限(Lv80)
    ('SVDF',45), #SV：DF上限(Lv80)
    #ーーーーーーーーーーーーーーーーーーーーーーーーー入力中
    ('Memo',63), #備考
]

forceInt = [
    'ST1', 'ST20', 'ST30', 'ST40', 'ST50', 'ST60', 'ST70',
    'SHMax', 'DRMax', 'PAMax', 'DFMax', 'MOfic', 'MStad', 'MClub', 
    'PHeig', 'PWeig', 'PBust', 'PWais', 'PHip',
    'SVSta', 'SVSH', 'SVDR', 'SVPA', 'SVDF'
]

PosList = [
    'LWG', 'CF', 'RWG',
    'LMF', 'CMF', 'RMF',
    'LSB', 'CB', 'RSB',
    'GK'
]

# Costume Data from Player Sheet (Basic Costume)
plCPos = [
    ('Cos',46),
    ('CosR',5),
    ('CosNote',37),
    ('CSkNam',48),
    ('CSkTyp',50),
    ('CSkCst',49),
    ('CSkMpy',51),
    ('CSkVCst',53),
    ('CSkVMpy',55)
]

def readPlayer( sh_player, players, costumes, pindex ):
    # read player and costume data from wear_sheet
    rows = sh_player.nrows
    c = 0

    for count in range(rows - 2):
        pl = {}
        cos0 = {}
        for i in range(len( plPPos )):
            pl[plPPos[i][0]] = sh_player.cell(count+2,plPPos[i][1]).value

        for i in range(len( plCPos )):
            cos0[plCPos[i][0]] = sh_player.cell(count+2,plCPos[i][1]).value

        if isnum( cos0['CosR'] ):
            ra = int(cos0['CosR'])
            if ( 0 < ra ) & ( ra < 7 ):
                rdif = (ra -1) * 25
                for prm in ['SH','DR','PA','DF']:
                    # print prm, pl[prm]
                    if isnum( pl[prm] ) & (pl[ prm ] > rdif):
                        pl[prm] = int( pl[prm] ) - rdif
                        for rr in range(5):
                            pl[(prm+'%d')%(rr+2)] = pl[prm]+25*(rr+1)
                    else:
                        pl[prm] = 0
                        for rr in range(5):
                            pl[(prm+'%d')%(rr+2)] = 0

            for x in forceInt:
                if isnum(pl[x]):
                    pl[x] = int( pl[x] )

            cos0['PName'] = pl['PName']
            # Basic costume parameters(+0)
            cos0['CosAST'] = 0
            cos0['CosASH'] = 0
            cos0['CosADR'] = 0
            cos0['CosAPA'] = 0
            cos0['CosADF'] = 0
            cos0['CosASum'] = 0
         
            costumes.append( cos0 )  #コスリストにビーナス(基本)ユニホーム追加
            players.append( pl )     #選手リストに選手データ追加

            # pindex[ 選手名 ] = [ 選手ID, [ コスチュームID, ... ]]
            pindex[pl['PName']] = [c, [c]]

            c += 1

    return( c )

# Costume Data from Wear Sheet (Special Costume)
weCPos = [
    ('PName',1),
    ('Cos',3),
    ('CosNote',5),
    ('CosR',6),
    ('CosAST',7),
    ('CosASH',8),
    ('CosADR',9),
    ('CosAPA',10),
    ('CosADF',11),
    ('CSkNam',13),
    ('CSkTyp',15),
    ('CSkCst',14),
    ('CSkMpy',16),
    ('CSkVCst',18),
    ('CSkVMpy',20)
]

def readWear( sh_wear, costumes, pindex ):
    # read costume data from wear_sheet
    rows = sh_wear.nrows
    c = 0

    for count in range(rows - 1):
        cos0 = {}
        for i in range( len( weCPos )):
            cos0[weCPos[i][0]] = sh_wear.cell(count+1,weCPos[i][1]).value
        if len(cos0['Cos']) > 0:
            sm = 0
            for x in ['ST','SH','DR','PA','DF']:
                if isnum( cos0['CosA'+x] ):
                    sm += int(cos0['CosA'+x])
                else:
                    sm = -999
            if sm >= 0:
                cos0['CosASum'] = sm
            else:
                cos0['CosASum'] = -1

            if cos0['PName'] in pindex:
                costumes.append( cos0 )
                pindex[cos0['PName']][1].append( len(costumes) -1 )
                c += 1

    return( c )


#
# Wiki Costume Table format
#
cosHdr = [
    u'//Char:%s',
    u'|CENTER:~名前|~レア度|>|>|>|>|CENTER:~ステータス加算値|>|>|CENTER:~スペシャルわざ|',
    u'|~|~|CENTER:~ｽﾀﾐﾅ|CENTER:~ｼｭｰﾄ|CENTER:~ﾄﾞﾘﾌﾞﾙ|CENTER:~ﾊﾟｽ|CENTER:~ﾃﾞｨﾌｪﾝｽ|CENTER:~名前|CENTER:~消費|CENTER:~効果|',
    u'|LEFT:150|LEFT:10|RIGHT:COLOR(green):45|RIGHT:COLOR(#aa0000):45|RIGHT:COLOR(#cc7711):45|RIGHT:COLOR(#bbaa11):45|RIGHT:COLOR(#1188aa):45|LEFT:160|RIGHT:30|LEFT:250|c'
]

cosOut = [
    # Cos,CosR,CosAST,CosASH,CosADR,CosAPA,CosADF,CSkNam,CSkCst,CSkTyp,CSkMpy
    u'|%s|☆%d|+%d|+%d|+%d|+%d|+%d|%s |%d|%sの効果を%s倍する|',
    u'|%s|☆%d|+%d|+%d|+%d|+%d|+%d|%s |  |%s |',
    u'|%s|☆%d|+  |+  |+  |+  |+  |%s |  |%s |',
    # CSkNam,CSkVCst,CSkTyp,CSkVMpy
    u'|~|~|~|~|~|~|~|%sV |%d|%sの効果を%s倍する|',
    u'|~|~|~|~|~|~|~|%sV |   | --- |'
]

def cosWikiPrint( players, costumes, pindex, enc ):
    c = 0
    for i in range(len(players)):
        PName = players[i]['PName']
        if dbg:
            print >>sys.stderr, pindex[PName]
            print >>sys.stderr, players[pindex[PName][0]]['PName']

        #Per Player Header
        for wrt in cosHdr:
            if wrt[0] == u'/':
                print (u'\n'+wrt%PName).encode(enc)
            else:
                print wrt.encode(enc)

        #Each Costume
        for x in pindex[PName][1]:
            c += 1
            if dbg:
                print >>sys.stderr, u' '+costumes[x]['Cos']

            if len( costumes[x]['CSkNam'] ) < 1:
                costumes[x]['CSkNam'] = ''

            if len( costumes[x]['CSkTyp'] ) < 1:
                costumes[x]['CSkTyp'] = ''
                costumes[x]['CSkCst'] = ''
                costumes[x]['CSkMpy'] = ''
                costumes[x]['CSkVCst'] = ''
                costumes[x]['CSkVMpy'] = ''

            if costumes[x]['CosASum'] < 0:
                line1 = cosOut[2]%(costumes[x]['Cos'],costumes[x]['CosR'],\
                                costumes[x]['CSkNam'],costumes[x]['CSkTyp'])
                line2 = cosOut[4]%(costumes[x]['CSkNam'])
                print line1.encode(enc)
                print line2.encode(enc)
                continue

            if isnum(costumes[x]['CSkCst']) & isnum(costumes[x]['CSkMpy']):
                line1 = cosOut[0]%(costumes[x]['Cos'],costumes[x]['CosR'],\
                                costumes[x]['CosAST'],costumes[x]['CosASH'],\
                                costumes[x]['CosADR'],costumes[x]['CosAPA'],\
                                costumes[x]['CosADF'],costumes[x]['CSkNam'],\
                                costumes[x]['CSkCst'],costumes[x]['CSkTyp'],\
                                fitos(costumes[x]['CSkMpy']))
            else:
                line1 = cosOut[1]%(costumes[x]['Cos'],costumes[x]['CosR'],\
                                costumes[x]['CosAST'],costumes[x]['CosASH'],\
                                costumes[x]['CosADR'],costumes[x]['CosAPA'],\
                                costumes[x]['CosADF'],costumes[x]['CSkNam'],\
                                costumes[x]['CSkTyp'])

            if isnum(costumes[x]['CSkVCst']) & isnum(costumes[x]['CSkVMpy']):
                line2 = cosOut[3]%(costumes[x]['CSkNam'],\
                                  costumes[x]['CSkVCst'],\
                                   costumes[x]['CSkTyp'],\
                                   fitos(costumes[x]['CSkVMpy']))
            else:
                line2 = cosOut[4]%(costumes[x]['CSkNam'] )

            print line1.encode(enc)
            print line2.encode(enc)

    return( c )


#
# all-costumes list CSV format
#

# version up
# ---- Costume plus value will display in MaxParam with [Base]+[Cos] format 
# ---- Position name will appear with category as FW/MF/DF/GK
# ---- Kana name field for sorting
coslistCsvHdr = u'Name,KanaName,PosC,Pos,' + \
    u'Costume,Rarity,PlusSum,' + \
    u'MaxST(Venus),MaxSH(V),MaxDR(V),MaxPA(V),MaxDF(V),' + \
    u'SkillType,Multiply,Cost,V-Multiply,V-Cost'

posC = {  u'LWG':u'1FW', u'CF':u'1FW', u'RWG':u'1FW',
          u'LMF':u'2MF', u'CMF':u'2MF', u'RMF':u'2MF',
          u'LSB':u'3DF', u'CB':u'3DF', u'RSB':u'3DF',
          u'GK':u'0GK'
         }

coslistWikiHdr = [
    u'|~名前|~ポジ|~衣装名|~ﾚｱ度|~Σ+|~体力&br;(ﾋﾞｰﾅｽ時)|~SH(V)|~DR(V)|~PA(V)|~DF(V)|~わざ種別|~倍率|~ｺｽﾄ|~倍率V|~ｺｽﾄV|',
    u'|LEFT:90|CENTER:30|LEFT:160|CENTER:20|CENTER:20 |CENTER:40|CENTER:COLOR(#aa0000):40|CENTER:COLOR(#cc7711):40|CENTER:COLOR(#bbaa11):40|CENTER:COLOR(#1188aa):40 |LEFT:60|RIGHT:40|RIGHT:30|RIGHT:COLOR(#bb0000):40|RIGHT:COLOR(#bb0000):30|c'
    ]

def cosListCsv( players, costumes, pindex, reffmt, enc ):
    spc = u',"%s"'
    spc1 = u'"%s"'
    spc2 = ' '

    if reffmt:
        print coslistWikiHdr[0].encode(enc)
        print coslistWikiHdr[1].encode(enc)
        print coslistCsvHdr.encode(enc)
        # spc = spc1 = u'|%s'
        spc2 = u'&br;'
    else:
        print coslistCsvHdr.encode(enc)

    c = 0

    # for each costume
    for i in range(len(costumes)):
        if len(costumes[i]['PName']) * len(costumes[i]['Cos']) == 0:
            continue

        pname = costumes[i]['PName']
        pid = pindex[pname][0]

        # 選手名 を バッファ u の先頭に設定
        if reffmt:
            dd = u'[['+pname+']]'
        else:
            dd = pname
        u = spc1%dd

        # かな名
        u += spc%players[pid]['PKana']

        # ポジション(縦位置、ポジション名)
        if players[pid]['Pos'] in PosList:
            u += spc%posC[ players[pid]['Pos'] ]
            dd = players[pid]['Pos']
            if reffmt:
                dd = u'[[' + dd + u']]'
            u += spc%dd
        else:
            u += spc%''+spc%players[pid]['Pos']

        # 衣装名,レア度
        u += spc%costumes[i]['Cos']
        if isnum( costumes[i]['CosR'] ):
            u += spc%(u'☆'+fitos(costumes[i]['CosR']))
        else:
            u += spc%''

        # 衣装加算値周りのリスト出力
        if costumes[i]['CosASum'] >= 0:
            # 衣装加算値が0以上であれば有効な衣装加算値が入っていると見なす
            u += spc%costumes[i]['CosASum']

            # スタミナ
            if costumes[i]['CosAST'] == 0:
                apd = u''
            else:
                apd = spc2+u'[+%d]'%costumes[i]['CosAST']
            if isnum( players[pid]['ST70'] ):
                dd = u'%d'%(players[pid]['ST70']+costumes[i]['CosAST'])+apd
            else:
                dd = u'---'+apd
            u += spc%dd
            # 能力値
            for x in ['SH','DR','PA','DF']:
                dt = costumes[i]['CosA'+x]
                if dt == 0:
                    apd = u''
                else:
                    apd = spc2+u'[+%d]'%dt
                if isnum( players[pid][x+'Max'] ):
                    dd = '%d'%(players[pid][x+'Max']+dt)+apd
                else:
                    dd = '---'+apd
                u += spc%dd
        else:
            u += spc%''+spc%'---'+spc%''+spc%''+spc%''+spc%''

        # スキル関係
        u += spc%costumes[i]['CSkTyp']
        u += spc%fitos( costumes[i]['CSkMpy'] )
        u += spc%fitos( costumes[i]['CSkCst'] )
        u += spc%fitos( costumes[i]['CSkVMpy'] )
        u += spc%fitos( costumes[i]['CSkVCst'] )

        #        if reffmt:
        #            u += u'|'

        print u.encode(enc)
        c += 1

    return c

# ----- to be modified with above Hdr & position category addition

def cosListCsvOld( players, costumes, pindex, reffmt, enc ):
    print u'Name,Costume,Rarity,Pos,'+\
          u'Stamina,Shoot,Dribble,Pass,Defence,PlusSum,'+\
          u'StaMax,ShMax,DrMax,PaMax,DefMax,'+\
          u'SkillType,Multiply,Cost,V-Multiply,V-Cost'
    c = 0

    for i in range(len(costumes)):
        if len(costumes[i]['PName']) * len(costumes[i]['Cos']) == 0:
            continue

        pname = costumes[i]['PName']
        pid = pindex[pname][0]
        
        if reffmt:
            u = u'"[['+pname+']]"'
        else:
            u = u'"'+pname+'"'

        u += u',"'+costumes[i]['Cos']+'"'
        if isnum( costumes[i]['CosR'] ):
            u += u',"☆%d"'%int(costumes[i]['CosR'])
        else:
            u += u',""'

        if players[pid]['Pos'] in PosList:
            if reffmt:
                u += ',"[[%s]]"'%players[pid]['Pos']
            else:
                u += ',"%s"'%players[pid]['Pos']
        else:
            u += u',""'

        if costumes[i]['CosASum'] >= 0:
            for x in ['ST','SH','DR','PA','DF']:
                u += u', +%d'%costumes[i]['CosA'+x]
            u += u', %d'%costumes[i]['CosASum']
            if isnum( players[pid]['ST70'] ):
                u += u',%d'%(players[pid]['ST70']+costumes[i]['CosAST'])
            else:
                u += ',"-"'
            for x in ['SH','DR','PA','DF']:
                if isnum( players[pid][x+'Max'] ):
                    u += ',%d'%(players[pid][x+'Max']+costumes[i]['CosA'+x])
                else:
                    u += ','
        else:
            u += u',"","","","","","","","","","",""'

        u += u',"'+costumes[i]['CSkTyp']+'"'
        u += u','+fitos( costumes[i]['CSkMpy'] )
        u += u','+fitos( costumes[i]['CSkCst'] )
        u += u','+fitos( costumes[i]['CSkVMpy'] )
        u += u','+fitos( costumes[i]['CSkVCst'] )

        print u.encode(enc)
        c += 1

    return c

wikiTmpl = [
    u"TITLE:%(PName)s",
    u"//コメント欄",
    u"// ★ 引継ぎ事項などあればこちらに ★",
    u"// ",
    u"#contents",
    u"#br",
    u"*概略 [#outline]",
    u"|CENTER:40|LEFT:120|CENTER:162|c",
    u"|>|CENTER:~プロフィール|~ビーナスユニフォーム|",
    u"|~名　前|%(PName)s|&attachref(./VU.jpg,zoom,162x288);|",
    u"// &attachref(./衣装2.jpg,zoom,162x288);|&attachref(./衣装3.jpg,zoom,162x288);|",
    u"|~あだ名|%(PHand)s|~|",
    u"|~身　長|%(PHeig)sｃｍ|~|",
    u"|~体　重|%(PWeig)sｋｇ  |~|",
    u"|~サイズ|Ｂ%(PBust)s　Ｗ%(PWais)s　Ｈ%(PHip)s|~|",
    u"|~誕生日|%(PBirt)s|~|",
    u"|~血液型|%(PBlod)s|~|",
    u"|~出身地|%(PLoca)s|~|",
    u"|~趣　味|%(PFavo)s|~|",
    u"|~ボイス|[[%(PCV)s]]|~|",
    u"//|CENTER:|CENTER:|CENTER:|CENTER:|CENTER:|c",
    u"//|>|~衣装４|~ |~ |~ |",
    u"//|>|&attachref(./衣装４.jpg,zoom,162x288);| | | |",
    u"**紹介 [#introduce]",
    u">",
    u"**評価 [#notes]",
    u">%(PNote)s",
    u"#br",
    u"*ステータス [#status]",
    u"//　　適性・経営",
    u"|CENTER:~ポジション|CENTER:~会社|CENTER:~ｽﾀｼﾞｱﾑ|CENTER:~ｸﾗﾌﾞﾊｳｽ|",
    u"|CENTER:70|CENTER:COLOR(orange):60|CENTER:COLOR(red):60|CENTER:COLOR(green):60|c",
    u"|[[%(Pos)s]]|%(MOfic)s|%(MStad)s|%(MClub)s|",
    u"//　　スタミナ",
    u"|~ |CENTER:~Lv1|CENTER:~Lv20 |CENTER:~Lv30 |CENTER:~Lv40 |CENTER:~Lv50 |CENTER:~Lv60 |CENTER:~Lv70 |",
    u"|~|~|CENTER:~☆1最大|CENTER:~☆2最大|CENTER:~☆3最大|CENTER:~☆4最大|CENTER:~☆5最大|CENTER:~☆6最大|",
    u"|LEFT:70|RIGHT:60|RIGHT:COLOR(blue):60|RIGHT:COLOR(blue):60|RIGHT:COLOR(blue):60|RIGHT:COLOR(blue):60|RIGHT:COLOR(blue):60|RIGHT:COLOR(blue):55|c",
    u"//       |初期|☆1 |☆2 |☆3 |☆4 |☆5 |☆6 |",
    u"//       |Lv1 |Lv20|Lv30|Lv40|Lv50|Lv60|Lv70|",
    u"|~ｽﾀﾐﾅ   | %(ST1)3s| %(ST20)3s| %(ST30)3s| %(ST40)3s| %(ST50)3s| %(ST60)3s| %(ST70)3s|",
    u"//　　能力",
    u"|~ |>|CENTER:~[[☆1]]|>|CENTER:~[[☆2]]|>|CENTER:~[[☆3]]|>|CENTER:~[[☆4]]|>|CENTER:~[[☆5]]|>|CENTER:~[[☆6]]|",
    u"|~|CENTER:~初期|CENTER:~最大|CENTER:~初期|CENTER:~最大|CENTER:~初期|CENTER:~最大|CENTER:~初期|CENTER:~最大|CENTER:~初期|CENTER:~最大|CENTER:~初期|CENTER:~最大|",
    u"|LEFT:70|RIGHT:30|RIGHT:COLOR(blue):30|RIGHT:30|RIGHT:COLOR(blue):30|RIGHT:30|RIGHT:COLOR(blue):30|RIGHT:30|RIGHT:COLOR(blue):30|RIGHT:30|RIGHT:COLOR(blue):30|RIGHT:30|RIGHT:COLOR(#bb4488):30|c",
    u"//       |☆1      |☆2      |☆3      |☆4      |☆5      |☆6      |",
    u"|~ｼｭｰﾄ   | %(SH)3s| 100| %(SH2)3s| 200| %(SH3)3s| 300| %(SH4)3s| 400| %(SH5)3s| 500| %(SH6)3s| %(SHMax)3s|",
    u"|~ﾄﾞﾘﾌﾞﾙ | %(DR)3s| 100| %(DR2)3s| 200| %(DR3)3s| 300| %(DR4)3s| 400| %(DR5)3s| 500| %(DR6)3s| %(DRMax)3s|",
    u"|~ﾊﾟｽ    | %(PA)3s| 100| %(PA2)3s| 200| %(PA3)3s| 300| %(PA4)3s| 400| %(PA5)3s| 500| %(PA6)3s| %(PAMax)3s|",
    u"|~ﾃﾞｨﾌｪﾝｽ| %(DF)3s| 100| %(DF2)3s| 200| %(DF3)3s| 300| %(DF4)3s| 400| %(DF5)3s| 500| %(DF6)3s| %(DFMax)3s|",
    u"//　　びびっどボード",
    u"|LEFT:70|CENTER:88|CENTER:88|CENTER:88|CENTER:88|CENTER:85|c",
    u"|~Sビーナス&br;最大値|~ｽﾀﾐﾅ|~ｼｭｰﾄ|~ﾄﾞﾘﾌﾞﾙ|~ﾊﾟｽ |~ﾃﾞｨﾌｪﾝｽ|",
    u"|~| %(SVSta)s&br;(ﾊﾟﾗﾒｰﾀｱｯﾌﾟ込み)|%(SVSH)3s|%(SVDR)3s|%(SVPA)3s|%(SVDF)3s|",
    u"|~固有スキル|>|>|LEFT:%(SVSkl)s|~コンバート先|CENTER:[[%(SVCvt)s]] |",
    u"|~スキル効果|>|>|>|>|LEFT:%(SVSkd)s|",
    u"//",
    u"*衣装 [#clothing]",
    u"#br",
    u"*コメント [#comments]",
    u"#pcomment(./コメント,50,reply)"
    ]

def wikibody( p ):
    l = 0
    for tpl in wikiTmpl:
        l += 1
        if ((50 <= l) & (l <= 55)) & (len(p['SVSkl']) == 0):
                continue
        print (tpl%p).encode(enc)


def playerWikiPrint( players, costumes, pindex, enc ):
    c = 0
    for i in range(len(players)):
        PName = players[i]['PName']
        if dbg:
            print >>sys.stderr, pindex[PName]
            print >>sys.stderr, players[pindex[PName][0]]['PName']

        #Per Player Header
        print (u'\n//%s ------------------------------'%PName).encode(enc)
        wikibody( players[i] )
        c += 1

    return c

chrTableHdr = [
    u"|CENTER:~[[名前>キャラ一覧/名前]]|CENTER:~[[ふりがな>キャラ一覧/ふりがな]]|CENTER:~[[ボイス>キャラ一覧/ボイス]]|CENTER:~[[適正>キャラ一覧/適正]]|CENTER:~☆|CENTER:~[[ST>キャラ一覧/スタミナ]]|CENTER:~[[SH>キャラ一覧/シュート]]|CENTER:~[[DR>キャラ一覧/ドリブル]]|CENTER:~[[PA>キャラ一覧/パス]]|CENTER:~[[DF>キャラ一覧/ディフェンス]]|CENTER:~[[会社>キャラ一覧/会社]]|CENTER:~[[ｽﾀｼﾞｱﾑ>キャラ一覧/スタジアム]]|CENTER:~[[ｸﾗﾌﾞﾊｳｽ>キャラ一覧/クラブハウス]]|CENTER:~[[誕生日>キャラ一覧/誕生日]]|CENTER:~備考|h",
    u"|LEFT:90|LEFT:110|LEFT:90|CENTER:50|RIGHT:COLOR(pink):20|RIGHT:COLOR(blue):30|RIGHT:COLOR(blue):30|RIGHT:COLOR(blue):30|RIGHT:COLOR(blue):30|RIGHT:COLOR(blue):30|RIGHT:COLOR(orange):20|RIGHT:COLOR(red):20|RIGHT:COLOR(green):20|CENTER:60|LEFT:|c"
    ]

chrTableFmt = [
    u"|[[%(PName)s]]|%(PKana)s|[[%(PCV)s]]|[[%(Pos)s]]|☆6| %(ST70)s|%(SHMax)s|%(DRMax)s|%(PAMax)s|%(DFMax)s|%(MOfic)s|%(MStad)s|%(MClub)s|%(PBirt)s|%(Memo)s|",
    u"|[[%(PName)s]]|%(PKana)s|[[%(PCV)s]]|[[%(Pos)s]]/[[%(SVCvt)s]]|☆7| %(SVSta)s|%(SVSH)s|%(SVDR)s|%(SVPA)s|%(SVDF)s|%(MOfic)s|%(MStad)s|%(MClub)s|%(PBirt)s|%(SVSkd)s&br;%(Memo)s|"
    ]

def playerChrIdxTblPrint( players, pindex, enc ):
    for t in chrTableHdr:
        print t.encode(enc)

    c = 0
    for i in range(len(players)):
        PName = players[i]['PName']
        if dbg:
            print >>sys.stderr, pindex[PName]
            print >>sys.stderr, players[pindex[PName][0]]['PName']

        if len(players[i]['SVCvt']) > 0:
            print (chrTableFmt[1]%players[i]).encode(enc)
        else:
            print (chrTableFmt[0]%players[i]).encode(enc)

        c += 1

    return c


if __name__ == '__main__':

    helpText = [
        u'Desc : Read びびび.xls and format',
        u' ---- control options',
        u'  -o filename  : Stdoutではなく、filenameに出力',
        u'  -i filename  : 入力ファイルを"びびび.xls"ではなくfilenameに変更',
        u'  -c encoding  : 出力エンコーディングを変更',
        u'  -u           : 出力エンコーディングをUTF-8に(既定はcp932)',
        u'  -r           : キャラ名リファレンス形式他をWiki参照に(更にL形式を一部変更)',
        u' ---- output format',
        u'  -L           : output costume list CSV',
        u'  -B           : output player data in Wiki body',
        u'  -T           : output player data in Wiki Char Index Table',
        u'  -W           : output costume list Wiki table',
        u' ---- other options',
        u'  -p           : デバグ出力',
        ]

    av = sys.argv
    ac = len(av)
    datafile = u'びびび.xls'
    enc = 'cp932'
    fmt = 'wiki'
    reffmt = 0

    outFileW = 'cos_tbl.txt'
    outFileL = 'skilllist.csv'
    outFileW = 'wikidata.txt'
    outfileT = 'chr_tbl.txt'

    if dbg: print "->", ac, av

    if ac > 1:
        i = 0
        while i < (ac-1):
            i = i+1
            flg = av[i]
            if dbg: print >>sys.stderr, "-->", i, "av[%d]"%i, flg

            if flg == '-o':
                try:
                    so = open( av[i+1], 'w' )
                except IOError, e:
                    print >>sys.stderr, 'File open error:',e
                    if dbg: print >>sys.stderr, 'FName =',av[i+1]
                    exit()
                sys.stdout = so
                i = i+1
            elif flg == '-i':
                try:
                    si = open( av[i+1], 'r' )
                except IOError, e:
                    print >>sys.stderr, 'File open error:',e
                    if dbg: print >>sys.stderr, 'FName =',av[i+1]
                    exit()
                si.close()
                datafile = av[i+1]
                i = i+1
            elif flg == '-u':
                enc = 'utf-8'
            elif flg == '-c':
                if len( av[i+1] ) > 0:
                    enc = av[i+1]
                    i = i+1
                else:
                    print >>sys.stderr, "-c <codec> ; no codec."
                    exit()
            elif flg == '-p':
                dbg = 1
            elif flg == '-r':
                reffmt = 1
            elif flg == '-L':
                fmt = 'csv'
            elif flg == '-B':
                fmt = 'body'
            elif flg == '-T':
                fmt = 'idxt'
            elif flg == '-W':
                fmt = 'wiki'
            elif flg == '-h':
                for j in range( len(helpText) ):
                    print >>sys.stderr, helpText[j].encode(enc)
                exit()

    try:
        book = xlrd.open_workbook( datafile )
    except IOError, e:
        print >>sys.stderr, 'File open error:',e,' File=',datafile
        exit()

    # 変数初期化
    players = []
    costumes = []
    pindex = {}

    # 'Player'シートの情報読出し
    try:
        sheet =  book.sheet_by_name('Player')
    except IOError, e:
        print >>sys.stderr, 'Sheet open error:',e,' ',datafile,'.Player'
        exit()
    readPlayer( sheet, players, costumes, pindex )

    # 'Wear'シートの情報読出し
    try:
        sh_wear =  book.sheet_by_name('Wear')
    except IOError, e:
        print >>sys.stderr, 'Sheet open error:',e,' ',datafile,'.Wear'
        exit()
    readWear( sh_wear, costumes, pindex )

    # 出力
    if fmt == 'csv':
        out = cosListCsv( players, costumes, pindex, reffmt, enc )
    elif fmt == 'body':
        out = playerWikiPrint( players, costumes, pindex, enc )
    elif fmt == 'wiki':
        out = cosWikiPrint( players, costumes, pindex, enc )
    elif fmt == 'idxt':
        out = playerChrIdxTblPrint( players, pindex, enc )

    print >>sys.stderr, 'Output',out,'lines'
