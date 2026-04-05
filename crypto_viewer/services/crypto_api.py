"""
CoinGecko API Service Module
Handles all interactions with the CoinGecko API
"""
import requests
import time
from functools import lru_cache
from config import Config


class CoinGeckoAPI:
    def __init__(self):
        self.base_url = Config.COINGECKO_API_BASE
        self.currency = Config.CURRENCY
        self._cache = {}
        self._cache_time = {}

    def _get_cached(self, key, fetch_func, timeout=None):
        """Simple cache implementation with timeout"""
        if timeout is None:
            timeout = Config.API_CACHE_TIMEOUT

        current_time = time.time()
        if key in self._cache and current_time - self._cache_time.get(key, 0) < timeout:
            return self._cache[key]

        data = fetch_func()
        self._cache[key] = data
        self._cache_time[key] = current_time
        return data

    def get_prices(self, coin_ids=None):
        """
        Get current prices for specified coins
        Returns dict with coin_id as key and price info as value
        """
        if coin_ids is None:
            coin_ids = [coin['id'] for coin in Config.SUPPORTED_COINS]

        ids_str = ','.join(coin_ids)
        cache_key = f'prices_{ids_str}'

        def fetch():
            url = f'{self.base_url}/simple/price'
            params = {
                'ids': ids_str,
                'vs_currencies': self.currency,
                'include_24hr_change': 'true',
                'include_market_cap': 'true',
                'include_24hr_vol': 'true'
            }
            try:
                response = requests.get(url, params=params, timeout=10)
                response.raise_for_status()
                return response.json()
            except requests.RequestException as e:
                print(f'Error fetching prices: {e}')
                return {}

        return self._get_cached(cache_key, fetch)

    def get_coin_list(self):
        """Get list of supported coins with current prices"""
        prices = self.get_prices()
        result = []

        for coin in Config.SUPPORTED_COINS:
            coin_id = coin['id']
            price_data = prices.get(coin_id, {})

            result.append({
                'id': coin_id,
                'symbol': coin['symbol'],
                'name': coin['name'],
                'price': price_data.get(self.currency, 0),
                'change_24h': price_data.get(f'{self.currency}_24h_change', 0),
                'market_cap': price_data.get(f'{self.currency}_market_cap', 0),
                'volume_24h': price_data.get(f'{self.currency}_24h_vol', 0)
            })

        return result

    def get_market_chart(self, coin_id, days='1'):
        """
        Get historical market data for a coin
        days: 1, 7, 30, 90, 365
        Returns price history data for chart
        """
        cache_key = f'chart_{coin_id}_{days}'

        def fetch():
            url = f'{self.base_url}/coins/{coin_id}/market_chart'
            params = {
                'vs_currency': self.currency,
                'days': days
            }
            try:
                response = requests.get(url, params=params, timeout=10)
                response.raise_for_status()
                data = response.json()

                # Format data for Chart.js
                prices = data.get('prices', [])
                return {
                    'labels': [p[0] for p in prices],  # timestamps
                    'prices': [p[1] for p in prices],  # prices
                }
            except requests.RequestException as e:
                print(f'Error fetching chart data: {e}')
                return {'labels': [], 'prices': []}

        # Longer cache for historical data
        timeout = 60 if days == '1' else 300
        return self._get_cached(cache_key, fetch, timeout)

    def get_coin_detail(self, coin_id):
        """Get detailed information about a specific coin"""
        cache_key = f'detail_{coin_id}'

        def fetch():
            url = f'{self.base_url}/coins/{coin_id}'
            params = {
                'localization': 'false',
                'tickers': 'false',
                'community_data': 'false',
                'developer_data': 'false'
            }
            try:
                response = requests.get(url, params=params, timeout=10)
                response.raise_for_status()
                data = response.json()

                market_data = data.get('market_data', {})
                return {
                    'id': coin_id,
                    'name': data.get('name', ''),
                    'symbol': data.get('symbol', '').upper(),
                    'image': data.get('image', {}).get('large', ''),
                    'price': market_data.get('current_price', {}).get(self.currency, 0),
                    'market_cap': market_data.get('market_cap', {}).get(self.currency, 0),
                    'market_cap_rank': market_data.get('market_cap_rank', 0),
                    'volume_24h': market_data.get('total_volume', {}).get(self.currency, 0),
                    'change_24h': market_data.get('price_change_percentage_24h', 0),
                    'change_7d': market_data.get('price_change_percentage_7d', 0),
                    'change_30d': market_data.get('price_change_percentage_30d', 0),
                    'high_24h': market_data.get('high_24h', {}).get(self.currency, 0),
                    'low_24h': market_data.get('low_24h', {}).get(self.currency, 0),
                    'ath': market_data.get('ath', {}).get(self.currency, 0),
                    'ath_date': market_data.get('ath_date', {}).get(self.currency, ''),
                    'circulating_supply': market_data.get('circulating_supply', 0),
                    'total_supply': market_data.get('total_supply', 0),
                }
            except requests.RequestException as e:
                print(f'Error fetching coin detail: {e}')
                return None

        return self._get_cached(cache_key, fetch, timeout=60)


# Global instance
crypto_api = CoinGeckoAPI()
