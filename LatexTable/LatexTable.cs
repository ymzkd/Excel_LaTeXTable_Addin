using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace LatexTable
{

    // CellSize
    public struct CellSize
    {
        public int col;
        public int row;
    }

    public enum Align { left, center, right };

    // そのうち doubleline, dot などもできたら
    public enum LineType { noline, line, undefine = -1 }

    public enum MergeType { lefttop, top, nottop, notmerged = -1 };

    /// <summary>
    /// Analysis Selected Range and Convert to minimal format
    /// </summary>    
    public class RangeConvert
    {

        /// <summary>
        /// 列数(↓方向,横列の数)
        /// </summary>
        public int rownum;

        /// <summary>
        /// 行数(→方向,縦列の数)
        /// </summary>
        public int colnum;

        /// <summary>
        /// テーブルの内部の文字列
        /// [rownum, colnum]
        /// </summary>
        public string[,] contents;

        /// <summary>
        /// 水平方向罫線
        /// [rownum+1, colnum]
        /// </summary>
        public LineType[,] hrule_map;

        /// <summary>
        /// 鉛直方向の罫線
        /// [rownum, colnum+1]
        /// </summary>
        public LineType[,] vrule_map;

        /// <summary>
        /// セルの文字寄せ
        /// [rownum, colnum]
        /// </summary>
        public Align[,] align_map;

        /// <summary>
        /// 結合セルであるか？結合セルなら左上・先頭列・その他かを調べて格納
        /// [rownum, colnum]
        /// </summary>
        public MergeType[,] merge_map;

        /// <summary>
        /// 結合を考慮した各セルのサイズの情報
        /// [rownum, colnum]
        /// </summary>
        public CellSize[,] cellsize_map;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="range">
        /// 整理するExcelのRange
        /// </param>
        /// <param name="hidden">
        /// Excel上で非表示のセルをTableでも非表示にするか
        /// </param>
        public RangeConvert(Excel.Range range, bool hidden = false)
        {

            List<int> row_index = new List<int>();
            List<int> col_index = new List<int>();

            // 列方向のインデックス作成
            for (int i = 1; i <= range.Rows.Count; i++)
            {
                Excel.Range line = range.Rows[i];
                //　表示/非表示考慮する設定でかつ非表示だったら
                if (hidden & line.Hidden)
                {
                    continue;
                }
                row_index.Add(i);
            }

            // 行方向のインデックス作成
            for (int i = 1; i <= range.Columns.Count; i++)
            {
                Excel.Range line = range.Columns[i];
                //　表示/非表示考慮する設定でかつ非表示だったら
                if (hidden & line.Hidden)
                {
                    continue;
                }
                col_index.Add(i);
            }


            // 初期化
            // Rownum, Colnum
            rownum = row_index.Count;
            colnum = col_index.Count;
            // Contents
            contents = new string[rownum, colnum];
            // Horizontal Rule Map
            hrule_map = new LineType[rownum + 1, colnum];
            // Vertical Rule Map
            vrule_map = new LineType[rownum, colnum + 1];
            // Align Map [rownum, colnum]
            align_map = new Align[rownum, colnum];
            // MergeMap [rownum, colnum]
            merge_map = new MergeType[rownum, colnum];
            // Cell Size Map [rownum, colnum]
            cellsize_map = new CellSize[rownum, colnum];


            // セルの文字寄せ,コンテンツ,マージ情報,マージサイズの取得
            for (int i = 0; i < row_index.Count; i++)
            {
                for (int j = 0; j < col_index.Count; j++)
                {
                    Excel.Range targetcell = range[row_index[i], col_index[j]];

                    // Align Map
                    align_map[i, j] = CellAlign(targetcell);

                    // マージの有無を確認
                    if (targetcell.MergeCells)
                    {
                        MergeType cell_mergeType = CheckMergeType(targetcell.MergeArea, targetcell, hidden);

                        // TopLeftだったら
                        if (cell_mergeType == MergeType.lefttop)
                        {
                            cellsize_map[i, j] = CheckCellSize(targetcell.MergeArea, hidden);
                            if (cellsize_map[i, j].col == 1 & cellsize_map[i, j].row == 1)
                            {
                                // TopLeftでもセルサイズが[1,1]だったらマージされてないと考える
                                merge_map[i, j] = MergeType.notmerged;
                            }
                            else
                            {
                                merge_map[i, j] = MergeType.lefttop;
                            }
                            contents[i, j] = targetcell.MergeArea[1, 1].Text;
                        }
                        // Topだったら
                        else if (cell_mergeType == MergeType.top)
                        {
                            cellsize_map[i, j] = new CellSize() { row = 1, col = 1 };
                            merge_map[i, j] = MergeType.top;
                            contents[i, j] = "";
                        }
                        // Top以外だったら
                        else
                        {
                            cellsize_map[i, j] = new CellSize() { row = 1, col = 1 };
                            merge_map[i, j] = MergeType.nottop;
                            contents[i, j] = "";
                        }

                    }
                    else
                    {
                        cellsize_map[i, j] = new CellSize() { row = 1, col = 1 };
                        merge_map[i, j] = MergeType.notmerged;
                        contents[i, j] = targetcell.Text;
                    }

                } // Column Loop
            } // Row Loop


            // 罫線情報の取得
            // 一番上の水平方向罫線の確認
            for (int j = 0; j < col_index.Count; j++)
            {
                hrule_map[0, j] = CheckLineType(range[row_index[0], col_index[j]], Excel.XlBordersIndex.xlEdgeTop);
            }

            // 左端の鉛直方向の罫線の確認
            for (int i = 0; i < row_index.Count; i++)
            {
                vrule_map[i, 0] = CheckLineType(range[row_index[i], col_index[0]], Excel.XlBordersIndex.xlEdgeLeft);
            }

            // 表の中身と右端と下の罫線の確認
            for (int i = 0; i < row_index.Count; i++)
            {

                for (int j = 0; j < col_index.Count; j++)
                {

                    // 簡単のために対象のセルを一旦変数に格納
                    Excel.Range targetcell = range[row_index[i], col_index[j]];


                    // 水平方向罫線
                    // Horizontal Rule Map -> 下側の罫線の確認
                    // このセルの下にもまだセルがある(一番下の列ではない)
                    if (i + 1 != row_index.Count)
                    {
                        // マージされていて、このセルの下のセルは先頭列ではない。
                        if (merge_map[i, j] != MergeType.notmerged & merge_map[i + 1, j] == MergeType.nottop)
                        {
                            hrule_map[i + 1, j] = LineType.noline;
                        }
                        else
                        {
                            hrule_map[i + 1, j] = CheckLineType(targetcell, Excel.XlBordersIndex.xlEdgeBottom);
                        }
                    }
                    else
                    {
                        hrule_map[i + 1, j] = CheckLineType(targetcell, Excel.XlBordersIndex.xlEdgeBottom);
                    }


                    // 鉛直方向罫線
                    // Vertical Rule Map -> 右側の罫線の確認
                    vrule_map[i, j + 1] = CheckLineType(targetcell, Excel.XlBordersIndex.xlEdgeRight);

                } // Column Loop

            } // Row Loop

        } // Constructor


        /// <summary>
        /// マージされたセルのサイズをチェック(セルの表示/非表示を考慮する)
        /// </summary>
        /// <param name="range">
        /// マージされた領域
        /// </param>
        /// <param name="hidden">
        /// Excel上で非表示のセルをTableでも非表示にするか
        /// </param>
        /// <returns>
        /// マージされたセルのサイズ
        /// </returns>
        private CellSize CheckCellSize(Excel.Range range, bool hidden = false)
        {
            CellSize cellsize = new CellSize() { col = 0, row = 0 };

            // 列方向のインデックス作成
            for (int i = 1; i <= range.Rows.Count; i++)
            {
                Excel.Range line = range.Rows[i];
                //　表示/非表示考慮する設定でかつ非表示だったら
                if (hidden & line.Hidden)
                {
                    continue;
                }
                cellsize.row++;
            }

            // 行方向のインデックス作成
            for (int i = 1; i <= range.Columns.Count; i++)
            {
                Excel.Range line = range.Columns[i];
                //　表示/非表示考慮する設定でかつ非表示だったら
                if (hidden & line.Hidden)
                {
                    continue;
                }
                cellsize.col++;
            }

            return cellsize;
        } // CheckCellSize


        /// <summary>
        /// Check Align and Convert to Align Type
        /// </summary>
        /// <param name="cell">
        /// 対象のセル
        /// </param>
        /// <returns>
        /// セルの文字寄せの状態
        /// </returns>
        private Align CellAlign(Excel.Range cell)
        {
            // Align Left
            if (cell.HorizontalAlignment == (int)Excel.XlHAlign.xlHAlignLeft)
            {
                return Align.left;
            }
            // Align Right
            else if (cell.HorizontalAlignment == (int)Excel.XlHAlign.xlHAlignRight)
            {
                return Align.right;
            }
            // Align Center
            else
            {
                return Align.center;
            }

        } // CellAlign

        /// <summary>
        /// MergeAreaに対して対象のセルがどの位置にあるか 
        /// </summary>
        /// <param name="merge_area">
        /// マージされた領域
        /// </param>
        /// <param name="check_cell">
        /// 対象のセル
        /// </param>
        /// <param name="hidden">
        /// Excel上で非表示のセルをTableでも非表示にするか
        /// </param>
        /// <returns>
        /// MergeType:マージされた領域における対象セルの位置関係
        /// </returns>
        private MergeType CheckMergeType(Excel.Range merge_area, Excel.Range check_cell, bool hidden = false)
        {

            int colhead = 1; int rowhead = 1;

            // 非表示セルを考慮する場合
            // 表示領域の先頭を見つける必要がある。
            if (hidden)
            {

                // 表示設定の最初の列を判定
                for (int i = 1; i <= merge_area.Rows.Count; i++)
                {
                    rowhead = i;
                    if (!merge_area.Rows[i].Hidden)
                    {
                        break;
                    }
                }

                // 表示設定の最初の行を判定
                for (int i = 1; i <= merge_area.Columns.Count; i++)
                {
                    colhead = i;
                    if (!merge_area.Columns[i].Hidden)
                    {
                        break;
                    }
                }

            }

            // 左上
            if (merge_area[rowhead, colhead].Address == check_cell.Address)
            {
                return MergeType.lefttop;
            }
            // 左上ではないが先頭列
            else if (merge_area[rowhead, colhead].Row == check_cell.Row)
            {
                return MergeType.top;
            }
            // それ以外
            else
            {
                return MergeType.nottop;
            }

        } // CheckMergeType

        // 対象のセルの指定位置の罫線の線種
        // 
        /// <summary>
        /// 対象のセルの指定位置の罫線の線種
        /// とりあえず現在は線の有無だけを判定
        /// </summary>
        /// <param name="cell">
        /// 対象のセル
        /// </param>
        /// <param name="pos">
        /// セルのどの位置(上下左右)の罫線を確認するか
        /// </param>
        /// <returns>
        /// 線種(現在は有無のみ)
        /// </returns>
        private LineType CheckLineType(Excel.Range cell, Excel.XlBordersIndex pos)
        {
            int t = cell.Borders[pos].LineStyle;

            // Line無し
            if (t == (int)Excel.XlLineStyle.xlLineStyleNone)
            {
                return LineType.noline;
            }
            else
            {
                return LineType.line;
            }
        } // CheckLineType

    } // Class RangeConvert

    /// <summary>
    /// Tabular要素の構成
    /// </summary>
    public class Tabular
    {
        string format;
        int colnum;
        int rownum;

        string indent = "  ";
        string space = " ";
        string separator = " & ";


        // tabular head
        private string tabular_head
        {
            get
            {
                return $"\\begin{{tabular}}{{{format}}}";
            }
        }

        // tabular foot
        private string tabular_foot
        {
            get
            {
                return "\\end{tabular}";
            }
        }

        /// <summary>
        /// 水平方向の罫線
        /// hrule_lines [rownum] 
        /// </summary>
        /// <example>
        /// hrule_lines[i] = "\cline{1-2} \cline{3-3}"
        /// </example>
        string hrule_top;
        string[] hrule_lines;

        // contents [rownum]
        List<string>[] contents;

        public Tabular(RangeConvert tb_dataset)
        {

            // 初期化
            colnum = tb_dataset.colnum;
            rownum = tb_dataset.rownum;
            // 内部のコンテンツリスト
            contents = new List<string>[rownum];
            for (int i = 0; i < tb_dataset.rownum; i++)
            {
                contents[i] = new List<string>();
            }
            // 列の水平罫線リスト
            hrule_lines = new string[rownum];
            LineType[] hrule_linetypes = new LineType[colnum];


            // 表の上の罫線をチェック
            for (int j = 0; j < colnum; j++)
            {
                hrule_linetypes[j] = tb_dataset.hrule_map[0, j];
            }
            hrule_top = hrule_str(hrule_linetypes);

            // フォーマットの作成
            LineType[] vrule_linetypes = new LineType[colnum + 1];
            Align[] valign_types = new Align[colnum];
            vrule_linetypes[0] = tb_dataset.vrule_map[0, 0];
            for (int j = 0; j < colnum; j++)
            {
                vrule_linetypes[j + 1] = tb_dataset.vrule_map[0, j + 1];
                valign_types[j] = tb_dataset.align_map[0, j];
            }
            format = tabular_format(vrule_linetypes, valign_types);


            // 中身のコンテンツ行を作成
            for (int i = 0; i < rownum; i++)
            {

                for (int j = 0; j < colnum; j++)
                {
                    // セルの水平方向罫線を生成するための配列を作る。
                    hrule_linetypes[j] = tb_dataset.hrule_map[i + 1, j];

                    // フォーマット
                    CellSize cell_size = tb_dataset.cellsize_map[i, j];

                    // 先頭列の文字装飾
                    LineType lvrule_top = tb_dataset.vrule_map[0, j]; // left vrule
                    LineType rvrule_top = tb_dataset.vrule_map[0, j + cell_size.col]; // right vrule +1をmulticol分増やす？
                    Align cell_align_top = tb_dataset.align_map[0, j];

                    // 文字装飾
                    LineType lvrule = tb_dataset.vrule_map[i, j]; // left vrule
                    LineType rvrule = tb_dataset.vrule_map[i, j + cell_size.col]; // right vrule +1をmulticol分増やす？
                    Align cell_align = tb_dataset.align_map[i, j];

                    MergeType cell_mergeType = tb_dataset.merge_map[i, j];
                    // コンテンツ
                    string cell_content = tb_dataset.contents[i, j];


                    // マージセルの左上
                    if (cell_mergeType == MergeType.lefttop)
                    {
                        contents[i].Add(multicell(cell_size, cell_content, lvrule, cell_align, rvrule, lvrule_top, cell_align_top, rvrule_top));
                    }
                    // マージセルの一番上の行 -> 一番左以外はまとめられるのでスキップ
                    else if (cell_mergeType == MergeType.top)
                    {
                        continue;
                    }
                    // マージセルの一番上以外の行
                    else if (cell_mergeType == MergeType.nottop)
                    {
                        // 基本結合されるので空白の行だが、罫線はmulticolumnなどで反映する必要がある場合もある。
                        contents[i].Add(multicell(cell_size, "", lvrule, cell_align, rvrule, lvrule_top, cell_align_top, rvrule_top));
                    }
                    // 結合セルではない。
                    else
                    {
                        contents[i].Add(multicell(cell_size, cell_content, lvrule, cell_align, rvrule, lvrule_top, cell_align_top, rvrule_top));
                    }
                }

                // セルの下水平方向罫線を出力
                hrule_lines[i] = hrule_str(hrule_linetypes);

            }

        }

        /// <summary>
        /// Tabularのフォーマットを作成
        /// </summary>
        /// <param name="vrule_lines">
        /// フォーマットのベースとなる鉛直方向罫線のリスト
        /// </param>
        /// <param name="valign">
        /// フォーマットのベースとなる各行の文字寄せ情報のリスト
        /// </param>
        /// <returns>
        /// Tabularフォーマット文字列
        /// -> "|c|c|l|c|"
        /// </returns>
        private string tabular_format(LineType[] vrule_lines, Align[] valign)
        {
            string s = "";
            s += linetype_to_vrule(vrule_lines[0]);
            for (int i = 0; i < valign.Length; i++)
            {
                s += align_to_str(valign[i]);
                s += linetype_to_vrule(vrule_lines[i + 1]);
            }

            return s;

        }

        /// <summary>
        /// セルサイズに応じて結合結果をLaTexの表現で返す。
        /// </summary>
        /// <param name="size">
        /// 結合セルサイズ
        /// </param>
        /// <param name="content">
        /// 結合セルの文字列
        /// </param>
        /// <param name="lvrule">
        /// 対象セルの左の鉛直罫線
        /// </param>
        /// <param name="align">
        /// 対象セルの文字寄せ
        /// </param>
        /// <param name="rvrule">
        /// 対象セルの右の鉛直罫線
        /// </param>
        /// <param name="lvrule_top">
        /// 対象セル行のフォーマットとなる左の鉛直罫線
        /// </param>
        /// <param name="align_top">
        /// 対象セル行のフォーマットとなる文字寄せ情報
        /// </param>
        /// <param name="rvrule_top">
        /// 対象セル行のフォーマットとなる右の鉛直罫線
        /// </param>
        /// <returns>
        /// 結合セルの結合結果
        /// </returns>
        private string multicell(
            CellSize size, string content,
            LineType lvrule, Align align, LineType rvrule,
            LineType lvrule_top, Align align_top, LineType rvrule_top)
        {
            string lvrule_str = linetype_to_vrule(lvrule);
            string rvrule_str = linetype_to_vrule(rvrule);
            string align_str = align_to_str(align);

            // 行と列の双方向結合
            if (size.col > 1 & size.row > 1)
            {
                return multicolumn(size.col, lvrule_str, align_str, rvrule_str,
                    multirow(size.row, content));
            }
            // 列(↓)方向のみの結合
            else if (size.row > 1)
            {
                // セルのフォーマットがデフォルトと異なる -> multicolumnで囲む
                if (lvrule_top != lvrule | align_top != align | rvrule_top != rvrule)
                {
                    return multicolumn(size.col, lvrule_str, align_str, rvrule_str, multirow(size.row, content));
                }
                // セルのフォーマットがデフォルトと同じ
                else
                {
                    return multirow(size.row, content);
                }
            }
            // 行(→)方向のみの結合
            else if (size.col > 1)
            {
                return multicolumn(size.col, lvrule_str, align_str, rvrule_str, content);
            }
            // 結合は無し
            else
            {
                // セルのフォーマットがデフォルトと異なる -> multicolumnで囲む
                if (lvrule_top != lvrule | align_top != align | rvrule_top != rvrule)
                {
                    return multicolumn(size.col, lvrule_str, align_str, rvrule_str, content);
                }
                // セルのフォーマットがデフォルトと同じ
                else
                {
                    return content;
                }
            }

        } // multicell

        // multicolumn format generator
        private string multicolumn(int colnum, string lrule, string align, string rrule, string content)
        {
            return $"\\multicolumn{{{colnum}}}{{{lrule}{align}{rrule}}}{{{content}}}";
        }

        // multirow format generator
        private string multirow(int rownum, string content)
        {
            return $"\\multirow{{{rownum}}}{{*}}{{{content}}}";
        }

        // multirow format generator
        private string cline(int start, int end)
        {
            return $"\\cline{{{start}-{end}}}";
        }


        // LineTypeをLaTexの縦方向の表現に変換
        private string linetype_to_vrule(LineType t)
        {
            if (t == LineType.line)
            {
                return "|";
            }
            else
            {
                return "";
            }
        } // linetype_to_vrule


        // AlignをLaTexの表現に変換
        private string align_to_str(Align a)
        {
            if (a == Align.left)
            {
                return "l";
            }
            else if (a == Align.center)
            {
                return "c";
            }
            else // Align Right
            {
                return "r";
            }
        } // align_to_str


        // 横方向の罫線の作成
        public string hrule_str(LineType[] hrule_list)
        {
            // 水平方向に全く罫線がない場合
            if (hrule_list.All(value => value == LineType.noline))
            {
                return "";
            }
            // 水平方向全体に罫線が引かれている場合
            else if (hrule_list.All(value => value == LineType.line))
            {
                return "\\hline";
            }
            // 水平方向に部分的に罫線が引かれている場合
            else
            {
                int start = 0; int end = 0;
                string s = "";
                while (start < hrule_list.Length)
                {
                    // 罫線のある始点を見つけた
                    if (hrule_list[start] == LineType.line)
                    {
                        end = start;

                        // 右のセルが範囲を超えない
                        if (end + 1 != hrule_list.Length)
                        {
                            // 次のセルに罫線がなかったら終了
                            while (hrule_list[end + 1] != LineType.noline)
                            {
                                end++;

                                // 次のセルが範囲から出るときは終了
                                if (end + 1 == hrule_list.Length)
                                {
                                    break;
                                }
                            }
                        }

                        s += cline(start + 1, end + 1);
                        start = end + 1;
                    }
                    else
                    {
                        start++;
                    }
                }
                return s;
            }
        }

        // tabularの中身のコンテンツを1行ごとにリストにする。-> contents line
        private List<string> contents_lines()
        {
            List<string> lines = new List<string>();
            // string row_str = string.Join(separator, tb_contents[i - 1]);
            for (int i = 0; i < rownum; i++)
            {
                string sep_contents = string.Join(separator, contents[i]);
                if (i == rownum - 1 & hrule_lines[i] == "")
                {
                    // 最終行で最後に水平方向の罫線がなかったら"//"を追加しない。
                    lines.Add(sep_contents);
                }
                else
                {
                    lines.Add(sep_contents + " \\\\ " + hrule_lines[i]);
                }
            }

            return lines;

        }

        /// <summary>
        /// 内部コンテンツのcontents lineをインデントしながらtabular環境に内挿
        /// </summary>
        /// <param name="contents_line">
        /// tabularの中身を行ごとにリストにしたもの
        /// </param>
        /// <returns>
        /// tabularを行ごとにリストにしたもの
        /// </returns>
        public List<string> Create_tabular()
        {

            string head = tabular_head + space + hrule_top;
            string foot = tabular_foot;

            List<string> tab_list = new List<string>();

            // Create tabular contents line
            tab_list.Add(head);
            foreach (string line in contents_lines())
            {
                // インデントを足してcontents lineを内部に展開
                tab_list.Add(indent + line);
            }
            tab_list.Add(foot);

            return tab_list;

        }


    } // Class LatexTabular

    /// <summary>
    /// LaTexのテーブル要素
    /// </summary>
    public class Table
    {
        /// <summary>
        /// \centeringによって表を中央に配置するか？
        /// </summary>
        public bool has_centering = false;

        /// <summary>
        /// 表の位置の制御
        /// </summary>
        public string position = "htpb";

        /// <summary>
        /// Captionを追加するか
        /// </summary>
        public bool has_caption = false;

        /// <summary>
        /// Captionの内容
        /// </summary>
        public string caption_content = "";

        /// <summary>
        /// Labelの追加するか
        /// </summary>
        public bool has_label = false;

        /// <summary>
        /// Labelの内容
        /// </summary>
        public string label_content = "";

        /// <summary>
        /// resizeboxによる本文の幅に合わせて表を拡大縮小するか
        /// </summary>
        public bool resize = false;

        private string indent = "  ";
        private string space = " ";

        private string table_head
        {
            get
            {
                return $"\\begin{{table}}[{position}]";
            }
        }

        private string table_foot = "\\end{table}";

        private string caption
        {
            get
            {
                return $"\\caption{{{caption_content}}}";
            }
        }

        private string label
        {
            get
            {
                return $"\\label{{{label_content}}}";
            }
        }

        private string centering = "\\centering";

        private string resize_head = "\\resizebox{\\textwidth}{!}{%";

        private string resize_foot = "}";

        public List<string> Create_table(List<string> contents_line)
        {

            List<string> tb_list = new List<string>();

            // Head
            tb_list.Add(table_head);

            // Add Centering
            if (has_centering)
            {
                tb_list.Add(indent + centering);
            }

            // Add Caption
            if (has_caption)
            {
                tb_list.Add(indent + caption);
            }

            // Add Label
            if (has_label)
            {
                tb_list.Add(indent + label);
            }

            // Resize Head
            if (resize)
            {
                tb_list.Add(indent + resize_head);
            }

            // Contents Lines
            foreach (string line in contents_line)
            {
                // インデントを足してcontents lineを内部に展開
                tb_list.Add(indent + line);
            }


            // Resize Foot
            if (resize)
            {
                tb_list.Add(indent + resize_foot);
            }

            // Foot
            tb_list.Add(table_foot);

            return tb_list;
        } // Create_table

        public List<string> Create_table(Tabular tab)
        {

            List<string> tb_list = new List<string>();

            // Head
            tb_list.Add(table_head);

            // Add Centering
            if (has_centering)
            {
                tb_list.Add(indent + centering);
            }

            // Add Caption
            if (has_caption)
            {
                tb_list.Add(indent + caption);
            }

            // Add Label
            if (has_label)
            {
                tb_list.Add(indent + label);
            }

            // Resize Head
            if (resize)
            {
                tb_list.Add(indent + resize_head);
            }

            // Contents Lines
            foreach (string line in tab.Create_tabular())
            {
                // インデントを足してcontents lineを内部に展開
                tb_list.Add(indent + line);
            }


            // Resize Foot
            if (resize)
            {
                tb_list.Add(indent + resize_foot);
            }

            // Foot
            tb_list.Add(table_foot);

            return tb_list;
        } // Create_table


    } // Class LaTexTable

}

