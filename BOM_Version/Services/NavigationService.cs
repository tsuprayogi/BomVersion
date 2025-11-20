using System;
using SAPbouiCOM;
using BOM_Version.Repositories;

namespace BOM_Version.Services
{
    public class NavigationService
    {
        private readonly DocumentRepository _repo;

        private string[] _keys = Array.Empty<string>();
        private int _index = -1;

        public NavigationService()
        {
            _repo = new DocumentRepository();
        }

        // ==========================================================
        // LOAD KEY LIST
        // ==========================================================
        public void LoadKeys()
        {
            string q = "SELECT Code FROM [@PRODBOM] ORDER BY Code";
            var rs = _repo.Exec(q);

            var list = new System.Collections.Generic.List<string>();

            while (!rs.EoF)
            {
                list.Add(rs.Fields.Item("Code").Value.ToString());
                rs.MoveNext();
            }

            _keys = list.ToArray();
            _index = (_keys.Length > 0 ? 0 : -1);
        }

        // ==========================================================
        // MOVE CURSOR
        // ==========================================================
        public string First()
        {
            if (_keys.Length == 0) return null;
            _index = 0;
            return _keys[_index];
        }

        public string Last()
        {
            if (_keys.Length == 0) return null;
            _index = _keys.Length - 1;
            return _keys[_index];
        }

        public string Next()
        {
            if (_keys.Length == 0) return null;
            if (_index < _keys.Length - 1)
                _index++;

            return _keys[_index];
        }

        public string Previous()
        {
            if (_keys.Length == 0) return null;
            if (_index > 0)
                _index--;

            return _keys[_index];
        }

        // ==========================================================
        // SET INDEX BY KEY
        // ==========================================================
        public void SetIndexByKey(string key)
        {
            if (_keys == null) return;

            int idx = Array.IndexOf(_keys, key);
            if (idx >= 0) _index = idx;
        }

        public string[] Keys => _keys;
    }
}
