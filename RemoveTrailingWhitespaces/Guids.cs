// Guids.cs
// MUST match guids.h
using System;

namespace Predelnik.RemoveTrailingWhitespaces
{
    static class GuidList
    {
        public const string guidRemoveTrailingWhitespacesPkgString = "70f718a3-a985-44e9-9e00-4c767c708ace";
        public const string guidRemoveTrailingWhitespacesCmdSetString = "9880ef45-cb7d-4531-bccf-d228fccbb119";

        public static readonly Guid guidRemoveTrailingWhitespacesCmdSet = new Guid(guidRemoveTrailingWhitespacesCmdSetString);
    };
}