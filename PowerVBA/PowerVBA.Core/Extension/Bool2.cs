namespace PowerVBA.Core.Extension
{
    public struct Bool2
    {
        public static Bool2 True = new Bool2(true);
        public static Bool2 False = new Bool2(false);

        public bool Value { get; set; }

        public Bool2(bool value)
        {
            this.Value = value;
        }

        public static implicit operator int(Bool2 value)
        {
            return value ? -1 : 0;
        }

        public static implicit operator Bool2(int state)
        {
            return state == -1;
        }

        public static implicit operator Bool2(bool value)
        {
            return new Bool2(value);
        }

        public static implicit operator bool(Bool2 bool2)
        {
            return bool2.Value;
        }
    }
}
