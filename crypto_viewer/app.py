"""
Crypto Price Viewer - Flask Application
A simple cryptocurrency price viewer with charts
"""
from flask import Flask, render_template, jsonify, request
from config import Config
from services.crypto_api import crypto_api

app = Flask(__name__)
app.config.from_object(Config)


@app.route('/')
def index():
    """Dashboard - Display price list"""
    coins = crypto_api.get_coin_list()
    return render_template('index.html',
                           coins=coins,
                           currency_symbol=Config.CURRENCY_SYMBOL)


@app.route('/chart/<coin_id>')
def chart(coin_id):
    """Chart detail page for a specific coin"""
    coin_detail = crypto_api.get_coin_detail(coin_id)
    if not coin_detail:
        return render_template('error.html', message='Coin not found'), 404

    return render_template('chart.html',
                           coin=coin_detail,
                           currency_symbol=Config.CURRENCY_SYMBOL)


@app.route('/api/prices')
def api_prices():
    """API endpoint for price data (for AJAX updates)"""
    coins = crypto_api.get_coin_list()
    return jsonify({
        'success': True,
        'data': coins,
        'currency': Config.CURRENCY,
        'currency_symbol': Config.CURRENCY_SYMBOL
    })


@app.route('/api/chart/<coin_id>')
def api_chart(coin_id):
    """API endpoint for chart data"""
    days = request.args.get('days', '1')
    if days not in ['1', '7', '30', '90', '365']:
        days = '1'

    chart_data = crypto_api.get_market_chart(coin_id, days)
    return jsonify({
        'success': True,
        'data': chart_data,
        'coin_id': coin_id,
        'days': days
    })


@app.route('/api/coin/<coin_id>')
def api_coin_detail(coin_id):
    """API endpoint for coin detail"""
    coin_detail = crypto_api.get_coin_detail(coin_id)
    if not coin_detail:
        return jsonify({'success': False, 'error': 'Coin not found'}), 404

    return jsonify({
        'success': True,
        'data': coin_detail
    })


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=Config.DEBUG)
