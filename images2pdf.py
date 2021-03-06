#!/usr/bin/env python3
# images2pdf
# 複数の画像ファイルを含むディレクトリーを、一つのPDFファイルに変換する
# 画像の縮小、右綴じの設定、画像のファイル名を利用した目次の生成に対応
# 対応する画像ファイルの拡張子: .tif .tiff .png .jpg .jpeg
# 画像ファイルに埋め込まれている解像度情報からページサイズを決定します
# 
# images2pdf.py input output [--pagelayout PAGELAYOUT] [--direction DIRECTION]
#                            [--resample RESAMPLE] [--outlines] [--linearize]
#                            [--objectstream] [--metafile]
# 
# sample:
# images2pdf.py input_dir output.pdf --pagelayout TwoPageRight --linearize --resample 150
# 
# [--pagelayout PAGELAYOUT] ページレイアウト（下記参照） / デフォルトは無指定（一般的なビューワーだとOneColumn）
# [--direction DIRECTION] 文字の方向（下記参照） / デフォルトは無指定（一般的なビューワーだとL2R）
# [--resample RESAMPLE] 拡張子がpng、jpg、またはjpegのファイルを縮小する際のDPI / デフォルト（0）は縮小しない
# [--outlines] 画像のファイル名から目次を生成するかどうか / デフォルト（無指定）は生成しない
# [--linearize] PDFをWeb用に最適化するかどうか / デフォルト（無指定）は最適化しない
# [--objectstream] PDF 1.7のオブジェクトストリームを使ってPDFを圧縮するかどうか / デフォルト（無指定）は圧縮しない
# [--metafile] 画像ファイル群と同じディレクトリーにあるメタファイル（下記参照）から設定を読み込むかどうか / デフォルト（無指定）は読み込まない
# 
# --pagelayout SinglePage -> 単一ページ表示
# --pagelayout OneColumn -> スクロールを有効にする
# --pagelayout TwoPageLeft -> 見開きページ表示
# --pagelayout TwoColumnLeft -> 見開きページでスクロール
# --pagelayout TwoPageRight -> 見開きページ表示 + 見開きページ表示で表紙を表示（表紙は単独表示）
# --pagelayout TwoColumnRight -> 見開きページでスクロール + 見開きページ表示で表紙を表示（表紙は単独表示）
# 
# --direction L2R -> 左綴じ（横書き）
# --direction R2L -> 右綴じ（縦書き）
# 
# メタファイルは、「@pdf --pagelayout TwoPageRight」などと、ここで指定できるオプションをファイル名にした空のファイルを、画像ファイルと同じディレクトリーに置いた上で、「images2pdf.py input_dir output.pdf --metafile」と--metafileオプションを付けてimages2pdf.pyを実行すると読み込まれます。
# 
# 目次はファイル名を次のようにして、--outlinesオプションを指定すると生成されます。⓿❶❷❸❹❺❻❼❽❾❿⓫⓬⓭⓮⓯が階層の指示に使えます。「※」以降の文字列は除外されます
# p.001 1 title ※note.jpg
# p.002 ❶1.1 title.jpg
# p.003 ❷1.1.1 title.jpg
# p.004.jpg
# p.005 2 title ※note ❶2.1 title.jpg
# p.006 ❶2.2 title ❶2.3 title.jpg
# p.007 3 title ⓿4 title.jpg
# p.008.jpg
# 
######### pikepdf、Imagemagickを使用しています。これらをインストールしないと動きません #########
# https://pikepdf.readthedocs.io/en/latest/installation.html
# https://imagemagick.org/script/download.php

import argparse
import re
import pikepdf
import pathlib
import tempfile
import subprocess

class Images2Pdf():
    def __init__(self, messageprint, dry_run=False):
        self.messageprint    = messageprint
        self.dry_run         = dry_run

        self.__default_pagelayout   = ''
        self.__default_direction    = ''
        self.__default_resample     = 0
        self.__default_outlines     = False
        self.__default_linearize    = False
        self.__default_objectstream = False
        self.__default_metafile     = False
    
    @property
    def default_pagelayout(self) -> str:
        return self.__default_pagelayout
 
    @default_pagelayout.setter
    def default_pagelayout(self, value: str):
        if type(value) is not str:
            raise TypeError(value)
        if value not in self.get_allow_pagelayout_list():
            raise ValueError(value)
        self.__default_pagelayout = value

    @staticmethod
    def get_allow_pagelayout_list() -> list:
        return ['', 'SinglePage', 'OneColumn', 'TwoColumnLeft', 'TwoColumnRight', 'TwoPageLeft', 'TwoPageRight']

    @property
    def default_direction(self) -> str:
        return self.__default_direction
 
    @default_direction.setter
    def default_direction(self, value: str):
        if type(value) is not str:
            raise TypeError(value)
        if value not in self.get_allow_direction_list():
            raise ValueError(value)
        self.__default_direction = value

    @staticmethod
    def get_allow_direction_list() -> list:
        return ['', 'L2R', 'R2L']

    @property
    def default_resample(self) -> int:
        return self.__default_resample
 
    @default_resample.setter
    def default_resample(self, value: int):
        if type(value) is not int:
            raise TypeError(value)
        self.__default_resample = value

    @property
    def default_outlines(self) -> bool:
        return self.__default_outlines
 
    @default_outlines.setter
    def default_outlines(self, value: bool):
        if type(value) is not bool:
            raise TypeError(value)
        self.__default_outlines = value
    
    @property
    def default_linearize(self) -> bool:
        return self.__default_linearize
 
    @default_linearize.setter
    def default_linearize(self, value: bool):
        if type(value) is not bool:
            raise TypeError(value)
        self.__default_linearize = value

    @property
    def default_objectstream(self) -> bool:
        return self.__default_objectstream
 
    @default_objectstream.setter
    def default_objectstream(self, value: bool):
        if type(value) is not bool:
            raise TypeError(value)
        self.__default_objectstream = value

    @property
    def default_metafile(self) -> bool:
        return self.__default_metafile
 
    @default_metafile.setter
    def default_metafile(self, value: bool):
        if type(value) is not bool:
            raise TypeError(value)
        self.__default_metafile = value

    # ArgumentParserを返す
    def get_argumentparser(self) -> argparse.ArgumentParser:
        parser = argparse.ArgumentParser()
        
        # 必須引数（メタファイルからは読み込まない）
        parser.add_argument('input')
        parser.add_argument('output')
        
        # オプション引数（メタファイルからも読み込む）
        self.__add_common_argumentparser_arguments(parser)

        # オプション引数（メタファイルからは読み込まない）
        if self.default_metafile:     parser.add_argument('--no-metafile'    , action='store_false')
        else:                         parser.add_argument('--metafile'       , action='store_true')
        
        return parser

    # ArgumentParserを返す（メタファイルから読み込むとき用）
    def get_metafile_argumentparser(self) -> argparse.ArgumentParser:
        parser = argparse.ArgumentParser()
        
        # オプション引数（メタファイルからも読み込む）
        self.__add_common_argumentparser_arguments(parser)
        
        return parser

    # ArgumentParser生成の共通部分
    def __add_common_argumentparser_arguments(self, parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
        # オプション引数（メタファイルからも読み込む）
        parser.add_argument('--pagelayout', default=self.default_pagelayout, choices=self.get_allow_pagelayout_list().remove(''))
        parser.add_argument('--direction' , default=self.default_direction , choices=self.get_allow_direction_list().remove('') )
        parser.add_argument('--resample', type=int, default=self.default_resample)
        if self.default_outlines:     parser.add_argument('--no-outlines'    , action='store_false')
        else:                         parser.add_argument('--outlines'       , action='store_true')
        if self.default_linearize:    parser.add_argument('--no-linearize'   , action='store_false')
        else:                         parser.add_argument('--linearize'      , action='store_true')
        if self.default_objectstream: parser.add_argument('--no-objectstream', action='store_false')
        else:                         parser.add_argument('--objectstream'   , action='store_true')
    
        return parser

    # メタファイルから読み取った引数をセットする
    def __set_metafile_args(self, args: argparse.Namespace):
        # オプション引数（メタファイルからも読み込む）
        self.pagelayout = args.pagelayout
        self.direction  = args.direction
        self.resample   = args.resample
        if self.default_outlines:     self.outlines     = not args.no_outlines
        else:                         self.outlines     =     args.outlines
        if self.default_linearize:    self.linearize    = not args.no_linearize
        else:                         self.linearize    =     args.linearize
        if self.default_objectstream: self.objectstream = not args.no_objectstream
        else:                         self.objectstream =     args.objectstream
    
    # 変換実行（argparse.Namespaceからオプションをセット）
    def set_args_and_convert(self, args: argparse.Namespace):
        # 必須引数（メタファイルからは読み込まない）
        self.source_dir_path = pathlib.Path(args.input)
        self.dest_file_path  = pathlib.Path(args.output)
        
        # オプション引数（メタファイルからも読み込む）
        self.__set_metafile_args(args)
        
        # オプション引数（メタファイルからは読み込まない）
        if self.default_metafile: self.metafile = not args.no_metafile
        else:                     self.metafile =     args.metafile

        self.__convert()

    # 変換実行（普通のメソッド引数でオプションをセット）
    def set_options_and_convert(self, source_dir_path: pathlib.PurePath, dest_file_path: pathlib.PurePath, pagelayout=None, direction=None, resample=None, outlines=None, linearize=None, objectstream=None, metafile=None):
        # 必須引数（メタファイルからは読み込まない）
        if not isinstance(source_dir_path, pathlib.PurePath): raise TypeError(source_dir_path)
        if not isinstance(dest_file_path , pathlib.PurePath): raise TypeError(dest_file_path)
        self.source_dir_path = source_dir_path
        self.dest_file_path  = dest_file_path

        # オプション引数（メタファイルからも読み込む）
        self.pagelayout   = pagelayout   if pagelayout   is not None else self.default_pagelayout
        self.direction    = direction    if direction    is not None else self.default_direction
        self.resample     = resample     if resample     is not None else self.default_resample
        self.outlines     = outlines     if outlines     is not None else self.default_outlines
        self.linearize    = linearize    if linearize    is not None else self.default_linearize
        self.objectstream = objectstream if objectstream is not None else self.default_objectstream
        
        # オプション引数（メタファイルからは読み込まない）
        self.metafile     = metafile     if metafile     is not None else self.default_metafile
        
        self.__convert()
    
    # 変換（内部メソッド）
    def __convert(self):
        # 画像ファイル群を取得
        image_files = []
        for image_file in self.source_dir_path.glob("*"):
            if image_file.is_file() and image_file.suffix.lower() in self.get_support_image_suffix_list():
                image_files.append(image_file)
        
        # メタデータ指示ファイル取得
        if self.metafile:
            parser = self.get_metafile_argumentparser()
            for meta_file_path in self.source_dir_path.glob("@pdf *"):
                args = parser.parse_args(meta_file_path.name[5:].split()) # "@pdf "の4文字の後を取り出し、前後のスペースを除去
                self.__set_metafile_args(args)
                break
        
        # 一時ディレクトリー作成
        with tempfile.TemporaryDirectory(dir=self.source_dir_path) as temp_dir:
            for image_file in image_files:
                self.create_page(image_file, temp_dir) # 個別PDF生成
            self.join_pdf(temp_dir) # 作成した個別PDFを連結

    # 個別PDF生成
    def create_page(self, image_file: pathlib.PurePath, temp_dir):
        # リサンプルする場合、モノクロ1bitのpng画像はグレースケールに変換する（jpegに1bitは存在しない）
        is_gray = 0 < self.resample and image_file.suffix.lower() == '.png' and ImagemagickWrapper.is_grayscale(image_file)
        # リサンプルする場合でもpngとjpeg限定
        is_resample = 0 < self.resample and image_file.suffix.lower() in ['.png', '.jpg', '.jpeg']
        
        # コマンド組み立て
        options = ['-path', str(temp_dir)]
        if is_gray: # グレースケール化が必要な場合は、オプションに追加
            options.extend(['-colorspace', 'gray'])
        if is_resample: # リサンプルが必要な場合は、オプションに追加
            options.extend(['-units', 'pixelspercentimeter', '-resample', round(float(self.resample) / 2.54, 2), '-unsharp', '2x2+0.7+0.02', '-compress', 'jpeg', '-quality', '60'])
        options.extend(['-format', 'pdf'])
        
        # コマンド実行（あくまで個別PDFのためメッセージは表示しない）
        ImagemagickWrapper.mogrify(options, image_file, temp_path=pathlib.Path(temp_dir), dry_run=self.dry_run)

    # PDF設定・連結
    def join_pdf(self, temp_dir):
        # 新PDF初期化
        pdf = pikepdf.new()
        version = pdf.pdf_version
        
        # PageLayout設定
        if 0 < len(self.pagelayout):
            if not hasattr(pdf.Root, 'PageLayout') \
                or pdf.Root.PageLayout != '/' + self.pagelayout:
                pdf.Root.PageLayout = pikepdf.Name('/' + self.pagelayout)
                if 'TwoPageLeft' == self.pagelayout or 'TwoPageRight' == self.pagelayout:
                    version = max(version, '1.5')
        
        # ViewerPreferences/Direction設定
        if 0 < len(self.direction):
            if not hasattr(pdf.Root, 'ViewerPreferences'):
                pdf.Root.ViewerPreferences = pikepdf.Dictionary()
            if not hasattr(pdf.Root.ViewerPreferences, 'Direction') \
                or pdf.Root.ViewerPreferences.Direction != '/' + self.direction:
                    pdf.Root.ViewerPreferences.Direction = pikepdf.Name('/' + self.direction)
                    version = max(version, '1.3')
        
        # 新PDFに個別PDFのページを追加していく
        page_path_list = sorted(list(pathlib.Path(temp_dir).glob("*.pdf")))
        for i, page_path in enumerate(page_path_list):
            page_pdf = pikepdf.open(page_path)
            version = max(version, page_pdf.pdf_version)
            pdf.pages.extend(page_pdf.pages)
        
        # outlines設定
        if self.outlines:
            new_outlines = Images2Pdf.generate_outlines(page_path_list)
            if 0 < len(new_outlines):
                with pdf.open_outline() as outline:
                    outline.root.extend(new_outlines)
    
        # 参照されていないオブジェクトを削除してから保存
        object_stream_mode = pikepdf.ObjectStreamMode.generate if self.objectstream else pikepdf.ObjectStreamMode.disable
        pdf.remove_unreferenced_resources()
        self.messageprint.print(normal='convert pdf: {}'.format(self.dest_file_path))
        if not self.dry_run:
            self.dest_file_path.parent.mkdir(parents=True, exist_ok=True) # ディレクトリーを作成してから
            pdf.save(self.dest_file_path, min_version=version, linearize=self.linearize, object_stream_mode=object_stream_mode)

    # 目次生成
    @staticmethod
    def generate_outlines(page_path_list):
        bookmarks_search_pattern = re.compile(r"[⓿❶❷❸❹❺❻❼❽❾❿⓫⓬⓭⓮⓯]?[^⓿❶❷❸❹❺❻❼❽❾❿⓫⓬⓭⓮⓯]+")
        bookmarks = []
        for page_number, page_path in enumerate(page_path_list):
            split_filenames = page_path.stem.split(maxsplit=1) # p.001 titletext.pdf
            if 2 == len(split_filenames) and 0 < len(split_filenames[1]):
                result = bookmarks_search_pattern.finditer(split_filenames[1])
                for r in result:
                    bookmark_name = r.group()
                    level = 0
                    for level_tmp, level_str in enumerate(['⓿', '❶', '❷', '❸', '❹', '❺', '❻', '❼', '❽', '❾', '❿', '⓫', '⓬', '⓭', '⓮', '⓯']):
                        if bookmark_name.startswith(level_str):
                            bookmark_name = bookmark_name[1:]
                            level = level_tmp
                            break
                    bookmark_name = bookmark_name.split('※', 1)[0]
                    bookmark_name = bookmark_name.strip()
                    if 0 < len(bookmark_name):
                        bookmarks.append({'level': level, 'name': bookmark_name, 'page_number': page_number})

        outlines = []
        if 0 == len(bookmarks):
            return outlines
        old_raw_level = None
        old_level = None
        old_item = None
        level_item_list = [ outlines ]
        for bookmark in bookmarks:
            item = pikepdf.OutlineItem(bookmark['name'], bookmark['page_number'])
            raw_level = bookmark['level']
            if old_raw_level is None:
                level = 0
            elif old_raw_level < raw_level:
                level = old_level + 1
                level_item_list.append(old_item.children)
            elif raw_level < old_raw_level:
                level = max(0, level - (old_raw_level - raw_level))
                if level < old_level:
                    del level_item_list[-(old_level - level):]
            else:
                level = old_level
                
            level_item_list[level].append(item)
            old_raw_level = raw_level
            old_level = level
            old_item = item
        return outlines

    # 対応する画像ファイルの拡張子
    @staticmethod
    def get_support_image_suffix_list() -> list:
        return ['.tif', '.tiff', '.png', '.jpg', '.jpeg']

class MessagePrint():
    def __init__(self, quiet=False, verbose=False, vv=False):
        if not type(quiet)   is bool: raise TypeError(quiet)
        self.__quiet   = quiet
        if not type(verbose) is bool: raise TypeError(verbose)
        self.__verbose = verbose
        if not type(vv) is bool: raise TypeError(vv)
        self.__vv = (verbose or vv)

    @property
    def quiet(self) -> bool: return self.__quiet

    @property
    def verbose(self) -> bool: return self.__verbose

    @property
    def vv(self) -> bool: return self.__vv

    def print(self, normal=None, verbose=None, vv=None):
        if self.quiet:
            return
        if self.vv:
            if vv is not None:
                print(vv)
                return
        if self.vv or self.verbose:
            if verbose is not None:
                print(verbose)
                return
        if normal is not None:
            print(normal)

class ImagemagickWrapper():
    @staticmethod
    def mogrify(options: list, path: pathlib.PurePath, temp_path=None, messageprint=None, message='mogrify', dry_run=False):
        cmd = ['/usr/bin/mogrify']
        cmd.extend(map(lambda x: str(x), options))
        if temp_path is not None:
            cmd.extend(['-define', 'registry:temporary-path=' + str(temp_path.resolve())])
        cmd.extend([str(path.resolve())])

        if messageprint is not None: messageprint.print(normal='{}: {}'.format(message, path), verbose='{}: {}'.format(message, ' '.join(cmd)))
        if not dry_run:
            subprocess.run(cmd, shell=False, check=True)

    @staticmethod
    def get_identify(format_list: list, path: pathlib.PurePath):
        format = '\n'.join(list(map(lambda x: '{0}\t%[{0}]'.format(x), format_list)))
        cmd = ['/usr/bin/identify', '-quiet', '-format', format, str(path.resolve())]
        stdout = subprocess.run(cmd, shell=False, stdout=subprocess.PIPE, text=True, check=True).stdout
        lines = stdout.splitlines()
        meta_data = {}
        for l in lines:
            a = l.split('\t', maxsplit=1)
            if 1 < len(a):
                meta_data[a[0]] = a[1]
            elif 0 < len(a):
                meta_data[a[0]] = ''
        return meta_data

    # グレースケール画像かどうか判定
    @staticmethod
    def is_grayscale(path: pathlib.PurePath) -> bool:
        meta_data = ImagemagickWrapper.get_identify(['r'], path)
        return 'r' in meta_data and 0 < len(meta_data['r']) and 'gray' in meta_data['r'].lower()

if __name__ == '__main__':
    images2pdf = Images2Pdf(MessagePrint())
    parser = images2pdf.get_argumentparser()
    images2pdf.set_args_and_convert(parser.parse_args())
