using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace HuanweiDzModels
{
    class LedgerBook : IEnumerable<LedgerItem>, IList<LedgerItem>
    {
        private LedgerItem[] _content = new LedgerItem[999];

        public LedgerItem this[int index]
        {
            get
            {
                return _content[index];
            }
            set
            {
                _content[index] = value;
            }
        }

        public int Count { get; private set; } = 0;

        public bool IsReadOnly => false;

        public void Add(LedgerItem item)
        {
            _content[Count] = item;
            Count++;
        }

        public void Clear()
        {
            _content = new LedgerItem[999];
            Count = 0;
        }

        public bool Contains(LedgerItem item)
        {
            for (int i = 0; i < Count; i++)
            {
                if (_content[i] == item) return true;
            }
            return false;
        }

        public void CopyTo(LedgerItem[] array, int arrayIndex)
        {
            throw new NotImplementedException();
        }

        public IEnumerator<LedgerItem> GetEnumerator()
        {
            for (int i = 0; i < Count; i++)
            {
                yield return _content[i];
            }
        }

        public int IndexOf(LedgerItem item)
        {
            for (int i = 0; i < Count; i++)
            {
                if (_content[i] == item) return i;
            }
            return -1;
        }

        public void Insert(int index, LedgerItem item)
        {
            //判定index是否在Count之内，比如Count = 9, Insert(10, item)需要报错
            if (index > Count) throw new IndexOutOfRangeException("插入序号大于序列大小。");
            for (int i = Count; i > index; i--) //从末尾开始全体右移一位直到插入处
            {
                _content[i] = _content[i - 1];

            }
            _content[index] = item;
            Count++;
        }

        public bool Remove(LedgerItem item)
        {
            //寻找到Item的位置
            int index = IndexOf(item);
            if (index == -1) //未找到时IndexOf返回-1；
            {
                return false;
            }
            RemoveAt(index);
            return true;
        }

        public void RemoveAt(int index)
        {
            for (int i = index; i < Count; i++)
            {
                _content[i] = _content[i + 1];

            }
            _content[Count - 1] = null;
            Count--;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            for (int i = 0; i < Count; i++)
            {
                yield return _content[i];
            }
        }
    }
}
