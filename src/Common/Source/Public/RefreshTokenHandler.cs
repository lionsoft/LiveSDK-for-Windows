using System;
using System.Threading.Tasks;

namespace Microsoft.Live
{
    public class RefreshTokenHandler : IRefreshTokenHandler
    {
        private RefreshTokenInfo _refreshTokenInfo;

        Task IRefreshTokenHandler.SaveRefreshTokenAsync(RefreshTokenInfo tokenInfo)
        {
            // Note: 
            // 1) In order to receive refresh token, wl.offline_access scope is needed.
            // 2) Alternatively, we can persist the refresh token.
            _refreshTokenInfo = tokenInfo;
            return Task.FromResult(0);
        }

        Task<RefreshTokenInfo> IRefreshTokenHandler.RetrieveRefreshTokenAsync()
        {
            return Task.FromResult(_refreshTokenInfo);
        }

        public RefreshTokenHandler()
        {
        }
        public RefreshTokenHandler(string refreshToken)
        {
            if (refreshToken == null) throw new ArgumentNullException("refreshToken");
            _refreshTokenInfo = new RefreshTokenInfo(refreshToken);
        }
    }
}