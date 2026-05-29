using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.IO.Compression;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
        private sealed class QuotaKey
        {
            public string TotalNo;
            public string ChapterSeq;
            public string OrderNo;

            public string Key
            {
                get { return TotalNo + "|" + ChapterSeq + "|" + OrderNo; }
            }
        }

        private static ToolStripMenuItem FindMenuItem(ToolStrip menu, string text)
        {
            foreach (ToolStripItem item in menu.Items)
            {
                ToolStripMenuItem menuItem = item as ToolStripMenuItem;
                if (menuItem != null && menuItem.Text == text)
                {
                    return menuItem;
                }
            }

            return null;
        }

        private static int FirstVisibleIndex(ToolStrip menu)
        {
            for (int i = 0; i < menu.Items.Count; i++)
            {
                if (menu.Items[i].Available && menu.Items[i].Visible)
                {
                    return i;
                }
            }

            return -1;
        }

        private static bool TryGetQuotaKey(DataGridViewRow row, out QuotaKey key)
        {
            key = null;
            if (row == null)
            {
                return false;
            }

            string totalNo = GetRowValue(row, "总概算序号de", "总概算序号");
            string chapterSeq = GetRowValue(row, "条目序号");
            string orderNo = GetRowValue(row, "顺号DE", "顺号");

            if (String.IsNullOrEmpty(totalNo) || String.IsNullOrEmpty(chapterSeq) || String.IsNullOrEmpty(orderNo))
            {
                Log("Selected row lacks quota key. total=" + totalNo + ", chapter=" + chapterSeq + ", order=" + orderNo);
                return false;
            }

            key = new QuotaKey { TotalNo = totalNo, ChapterSeq = chapterSeq, OrderNo = orderNo };
            return true;
        }

        private static string GetRowValue(DataGridViewRow row, params string[] names)
        {
            DataRowView rowView = row.DataBoundItem as DataRowView;
            if (rowView != null)
            {
                foreach (string name in names)
                {
                    if (rowView.DataView.Table.Columns.Contains(name))
                    {
                        object value = rowView[name];
                        if (value != null && value != DBNull.Value)
                        {
                            return Convert.ToString(value).Trim();
                        }
                    }
                }
            }

            if (row.DataGridView != null)
            {
                foreach (string name in names)
                {
                    if (row.DataGridView.Columns.Contains(name))
                    {
                        object value = row.Cells[name].Value;
                        if (value != null)
                        {
                            return Convert.ToString(value).Trim();
                        }
                    }
                }
            }

            return null;
        }
    }
}
