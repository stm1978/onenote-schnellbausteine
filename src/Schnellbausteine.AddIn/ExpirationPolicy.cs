using System;

namespace Schnellbausteine.AddIn
{
    internal enum BetaStatus
    {
        Active,
        Warning,
        Expired
    }

    internal interface ILocalClock
    {
        DateTimeOffset LocalNow { get; }
    }

    internal sealed class SystemLocalClock : ILocalClock
    {
        public DateTimeOffset LocalNow
        {
            get { return DateTimeOffset.Now; }
        }
    }

    internal sealed class ExpirationPolicy
    {
        internal static readonly DateTime WarningStarts =
            new DateTime(2026, 12, 1, 0, 0, 0, DateTimeKind.Unspecified);
        internal static readonly DateTime ExpirationStarts =
            new DateTime(2027, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);

        private readonly ILocalClock clock;

        public ExpirationPolicy(ILocalClock clock)
        {
            if (clock == null) throw new ArgumentNullException("clock");
            this.clock = clock;
        }

        public BetaStatus CurrentStatus
        {
            get
            {
                DateTime localWallTime = clock.LocalNow.DateTime;
                if (localWallTime >= ExpirationStarts) return BetaStatus.Expired;
                if (localWallTime >= WarningStarts) return BetaStatus.Warning;
                return BetaStatus.Active;
            }
        }

        public bool AllowsFeatures
        {
            get { return CurrentStatus != BetaStatus.Expired; }
        }
    }
}
