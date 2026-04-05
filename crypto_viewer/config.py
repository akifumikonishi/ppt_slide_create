"""
Crypto Viewer Configuration
"""
import os

class Config:
    # Flask settings
    SECRET_KEY = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-production')
    DEBUG = os.environ.get('FLASK_DEBUG', 'True').lower() == 'true'

    # CoinGecko API settings
    COINGECKO_API_BASE = 'https://api.coingecko.com/api/v3'

    # Supported cryptocurrencies
    SUPPORTED_COINS = [
        {'id': 'bitcoin', 'symbol': 'BTC', 'name': 'Bitcoin'},
        {'id': 'ethereum', 'symbol': 'ETH', 'name': 'Ethereum'},
        {'id': 'ripple', 'symbol': 'XRP', 'name': 'XRP'},
        {'id': 'cardano', 'symbol': 'ADA', 'name': 'Cardano'},
        {'id': 'solana', 'symbol': 'SOL', 'name': 'Solana'},
        {'id': 'dogecoin', 'symbol': 'DOGE', 'name': 'Dogecoin'},
        {'id': 'polkadot', 'symbol': 'DOT', 'name': 'Polkadot'},
        {'id': 'litecoin', 'symbol': 'LTC', 'name': 'Litecoin'},
    ]

    # API rate limiting
    API_CACHE_TIMEOUT = 30  # seconds

    # Currency for price display
    CURRENCY = 'jpy'
    CURRENCY_SYMBOL = '¥'
