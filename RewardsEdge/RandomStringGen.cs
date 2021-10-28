using System;
using System.Linq;

namespace RewardsEdge
{
    class RandomStringGen
    {
        Random r;
        string chars;
        public RandomStringGen()
        {
            r = new Random();
            chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
        }

        public string GenString(int length)
        {
            return new string(Enumerable.Repeat(chars, length).Select(s => s[r.Next(s.Length)]).ToArray());
        }
    }
}
