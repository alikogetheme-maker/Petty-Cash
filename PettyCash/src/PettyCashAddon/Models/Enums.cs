namespace PettyCashAddon.Models
{
    public enum Shift
    {
        Matin,
        ApresMidi,
        Soir
    }

    public enum Direction
    {
        Recette,
        Depense
    }

    public enum SessionStatus
    {
        Open,
        Closed
    }

    internal static class EnumCodes
    {
        public static string ToCode(Shift shift)
        {
            switch (shift)
            {
                case Shift.Matin: return "M";
                case Shift.ApresMidi: return "A";
                case Shift.Soir: return "S";
                default: return "M";
            }
        }

        public static Shift ShiftFromCode(string code)
        {
            switch (code)
            {
                case "A": return Shift.ApresMidi;
                case "S": return Shift.Soir;
                default: return Shift.Matin;
            }
        }

        public static string ToCode(Direction direction)
        {
            return direction == Direction.Recette ? "R" : "D";
        }

        public static Direction DirectionFromCode(string code)
        {
            return code == "R" ? Direction.Recette : Direction.Depense;
        }

        public static string ToCode(SessionStatus status)
        {
            return status == SessionStatus.Open ? "O" : "C";
        }

        public static SessionStatus StatusFromCode(string code)
        {
            return code == "O" ? SessionStatus.Open : SessionStatus.Closed;
        }
    }
}
