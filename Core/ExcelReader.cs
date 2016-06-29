using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Core
{
    public class ExcelReader
    {
        private readonly DbCommander _cmd;
        private readonly ExcelConfiguration _config;
        private readonly int _step;

        public ExcelReader(DbCommander cmd, ExcelConfiguration config)
        {
            _cmd = cmd;
            _config = config;
            _step = 6;
        }

        public void Read(string file)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(file))
                {
                    pck.Load(stream);
                }
                var wss = pck.Workbook.Worksheets.ToArray();
                var agencyId = new FileInfo(file).Name.Split('_')[0];
                for (var i = 1; i < 3; i++)
                {
                    var ws = wss[i];
                    for (var j = 3; j <= 333; j += _step)
                    {
                        var categoryName = ((string)ws.Cells["A" + j].Value).Split('(')[0].Trim();
                        foreach (var d in _config.Discipline)
                        {
                            double value;
                            var x = ws.Cells[d.Value + j].Value == null ? "" : ws.Cells[d.Value + j].Value.ToString();
                            if (double.TryParse(x, out value) && Math.Abs(value) > 0)
                            {
                                var disciplineId = _cmd.InsertDiscipline(agencyId, d.Key, categoryName, i, value);

                                #region SubType

                                var row = 2;
                                var subType = new Dictionary<string, string>();
                                switch (d.Key)
                                {
                                    case "Display":
                                        subType = _config.Display;
                                        break;

                                    case "VDO":
                                        subType = _config.Vdo;
                                        break;

                                    case "YouTube":
                                        row = 5;
                                        subType = _config.YouTube;
                                        break;

                                    case "Facebook":
                                        row = 5;
                                        subType = _config.Facebook;
                                        break;

                                    case "Instagram":
                                        subType = _config.Instagram;
                                        break;

                                    case "Creative":
                                        subType = _config.Creative;
                                        break;
                                }

                                #endregion SubType

                                foreach (var s in subType)
                                {
                                    double percent;
                                    var y = ws.Cells[s.Value + (j + row)].Value == null
                                        ? ""
                                        : ws.Cells[s.Value + (j + row)].Value.ToString();
                                    if (double.TryParse(y, out percent) && Math.Abs(percent) > 0)
                                    {
                                        _cmd.InsertSubDiscipline(disciplineId, s.Key, percent);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}