using System;
using System.Diagnostics;

namespace ODIS.ODIS
{
    class Hyperlink
    {
        string _text;
        string _link;
        bool _visited;

        public Hyperlink(string text, string link)
        {
            _text = text;
            _link = link;
            _visited = false;
        }
        public bool Visited
        {
            get { return _visited; }
        }
        public Process Activate()
        {
            _visited = true;
            return Process.Start(_link);
        }
        override public string ToString()
        {
            return _text;
        }

        public string Link
        {
            get { return _link; }
        }
    }
}
