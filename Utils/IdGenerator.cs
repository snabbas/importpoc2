
namespace ImportPOC2.Utils
{
    static class IdGenerator
    {
        private static long _currentId = 0;

        public static long getNextid()
        {
            return --_currentId;
        }

        public static void resetIds()
        {
            _currentId = 0;
        }
    }
}
