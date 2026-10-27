using System;
using System.Collections.Generic;

namespace Daisy.SaveAsDAISY.Conversion.Update
{
    /// <summary>
    /// Update channels available to the user
    /// </summary>
    public static class UpdateChannel
    {
        public enum Enum
        {
            /// <summary>
            /// Only stable releases are proposed
            /// </summary>
            Stable,
            /// <summary>
            /// Stable and pre-releases are proposed
            /// </summary>
            Beta
        }

        public static readonly Dictionary<Enum, string> Values = new Dictionary<Enum, string>()
        {
            { Enum.Stable, "stable" },
            { Enum.Beta, "beta" },
        };

        public static readonly Dictionary<string, Enum> Keys = new Dictionary<string, Enum>()
        {
            { "stable", Enum.Stable },
            { "beta", Enum.Beta },
        };
    }
}
