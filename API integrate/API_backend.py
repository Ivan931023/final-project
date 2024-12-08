import os
import io
from datetime import datetime
from dateutil.relativedelta import relativedelta
from flask import Flask, request, jsonify, send_file
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import xlsxwriter
from talib import abstract
import openai
from google.cloud import vision


app = Flask(__name__)


openai.api_key = ""


if not os.path.exists('images'):
    os.makedirs('images')

if not os.path.exists('excels'):
    os.makedirs('excels')
    
@app.route('/api/analyze_data', methods=['POST'])
def analyze_data():
    try:
        data = request.json
        stock = data.get('stock')
        start = data.get('start')
        end = data.get('end')
        show_sma = False
        
        if not all([stock, start, end]):
            return jsonify({'error': '缺少必需的參數'}), 400

        result = []
        workbook_path = os.path.join('excels', f'{stock}_data.xlsx')
        workbook = xlsxwriter.Workbook(workbook_path)
        worksheet = workbook.add_worksheet(stock)
        pre = datetime.strptime(start, '%Y-%m-%d') - relativedelta(months=2)
        pre = pre.strftime('%Y-%m-%d')

        df = yf.Ticker(stock).history(start=pre, end=end)
        df.index = df.index.tz_localize(None)
        cdf = df['Close']

        def draw_pic(inp):
            plt.figure(figsize=(6, 3))
            plt.plot(inp['x_val'], inp['y_val'][0], label= 'close', marker='.',markersize= 1,c = 'k')
            if show_sma:
                plt.plot(inp['x_val'], inp['y_val'][1], label= '5SMA', marker='.',markersize= 0,c = 'orange')
                plt.plot(inp['x_val'], inp['y_val'][2], label= '20SMA', marker='.',markersize= 0,c = 'green')
            plt.title(f'{stock} {inp['name']}') 
            plt.xlabel('Date')
            plt.ylabel(f'{inp['ylabel']}')
            plt.grid()
            plt.legend()
            plt.xticks(rotation=45)
            plt.tight_layout()
            plt.savefig(os.path.join('images', f'{stock} {inp['name']}.png')) 
            plt.close()  
            worksheet.insert_image(inp['loc'], os.path.join('images', f'{stock} {inp['name']}.png'))
            result.append(os.path.join('images', f'{stock} {inp['name']}.png'))
        
        def draw_pic_tec(inp):
            plt.figure(figsize=(6, 3))
            plt.plot(inp['x_val'], inp['y_val'][0], label = inp['lab'][0], marker='.',markersize= 0,color = 'blue')
            plt.plot(inp['x_val'], inp['y_val'][1], label = inp['lab'][1], marker='.',markersize= 0,color = 'hotpink')
            plt.bar(inp['x_val'], inp['y_val'][2], label = inp['lab'][2],color = ['green' if value > 0 else 'red' for value in inp['y_val'][2]])
            plt.legend()
            plt.title(f'{stock} {inp['name']}') 
            plt.xlabel('Date')
            plt.ylabel(f'{inp['ylabel']}')
            plt.grid()
            plt.xticks(rotation=45)
            plt.tight_layout()
            plt.savefig(os.path.join('images', f'{stock} {inp['name']}.png'))
            plt.close() 
            worksheet.insert_image(inp['loc'], os.path.join('images', f'{stock} {inp['name']}.png'))
            result.append(os.path.join('images', f'{stock} {inp['name']}.png'))
            
        def draw_pic_totval(inp):
            plt.figure(figsize=(10, 3))
            plt.plot([inp.index[0], inp.index[-1]], [100,100], label= '100%', ls = '--',linewidth = '5',c = 'red')
            plt.plot(inp.index, inp['st0'], label= 'buy and hold', marker='.',markersize= 0,c = 'navy')
            plt.plot(inp.index, inp['st1'], label= 'st1', marker='.',markersize= 0,c = 'brown')
            plt.plot(inp.index, inp['st2'], label= 'st2', marker='.',markersize= 0,c = 'cyan')
            plt.plot(inp.index, inp['st_sma5'], label= 'sma5', marker='.',markersize= 0,c = 'lime')
            plt.plot(inp.index, inp['st_macd'], label= 'macd', marker='.',markersize= 0,c = 'gray')
            plt.plot(inp.index, inp['st_kd'], label= 'kd', marker='.',markersize= 0,c = 'orange')
            plt.plot(inp.index, inp['st_rsi'], label= 'rsi', marker='.',markersize= 0,c = 'magenta')
            plt.title(f'{stock} earning rate') 
            plt.xlabel('date')
            plt.ylabel('%')
            plt.grid()
            plt.legend()
            plt.xticks(rotation=45)
            plt.tight_layout()
            plt.savefig(os.path.join('images', f'{stock} earning rate.png')) 
            plt.close()  
            worksheet.insert_image('U1', os.path.join('images', f'{stock} earning rate.png'))
            result.append(os.path.join('images', f'{stock} earning rate.png'))
        

        df.rename(columns={'High': 'high', 'Low': 'low', 'Close': 'close'}, inplace=True)
        sma20_df = abstract.SMA(df, 20)
        sma5_df = abstract.SMA(df, 5)
        macd_df = abstract.MACD(df, fastperiod=12, slowperiod=26, signalperiod=9)
        kd_df = abstract.STOCH(df, fastk_period=9, slowk_period=5, slowd_period=5)
        rsif_df = abstract.RSI(df, 5)
        rsis_df = abstract.RSI(df, 10)
        
        sma20_df.name = 'sma20'
        sma5_df.name = 'sma5'
        rsif_df.name = 'rsif'
        rsis_df.name = 'rsis'
        
        res = pd.concat([cdf, sma20_df, sma5_df, macd_df, kd_df, rsif_df, rsis_df], axis=1)
        res = res.loc[res.index >= start]

        worksheet.write('A1', 'Date')
        worksheet.write('B1', 'Close Price')
        for idx in range(len(res)):
            worksheet.write(idx + 1, 0, res.index[idx].strftime('%Y-%m-%d')) 
            worksheet.write(idx + 1, 1, float(round(res['Close'].iloc[idx], 2)))

        draw_pic({
        'x_val': res.index,
        'y_val': [res['Close'],res['sma5'],res['sma20']],
        'name' : 'Closing Prices History',
        'ylabel': 'price',
        'loc': 'C1'
        })
        draw_pic_tec({
            'x_val': res.index,
            'y_val': [res['macd'], res['macdsignal'], res['macdhist']],
            'name' : 'MACD History',
            'ylabel': 'unit',
            'loc': 'C31',
            'lab': ['DIF', 'MACD', 'OSC'],
        })
        draw_pic_tec({
            'x_val': res.index,
            'y_val': [res['slowk'], res['slowd'], res['slowk'] - res['slowd']],
            'name' : 'KD History',
            'ylabel': 'unit',
            'loc': 'C61',
            'lab': ['K', 'D', 'K-D'],
        })
        draw_pic_tec({
            'x_val': res.index,
            'y_val': [res['rsif'], res['rsis'], res['rsif'] - res['rsis']],
            'name' : 'RSI History',
            'ylabel': 'unit',
            'loc': 'C91',
            'lab': ['RSI(5)', 'RSI(10)', 'RSI(5)-RSI(10)'],
        })
    
    
        def strategy0(dframe):
            rate_data = pd.Series(dframe['Close'] / dframe['Close'].iloc[0]*100, name = 'st0')
            return rate_data
        def strategy1(dframe):
            rate = 100
            rate_data = [rate]  
            hold = 1
            for i in range(1,len(dframe['Close'])):
                if hold:
                    rate = rate*dframe['Close'].iloc[i]/dframe['Close'].iloc[i-1]
                if dframe['Close'].iloc[i-1] <= dframe['Close'].iloc[i]:
                    hold = 1
                else: 
                    hold = 0
                rate_data.append(rate)
            rate_data = pd.Series(rate_data,index = dframe.index, name = 'st1')
            return rate_data
        def strategy2(dframe):
            rate = 100
            rate_data = [rate]  
            hold = 1
            for i in range(1,len(dframe['Close'])):
                if hold:
                    rate = rate*dframe['Close'].iloc[i]/dframe['Close'].iloc[i-1]
                if dframe['Close'].iloc[i-1] <= dframe['Close'].iloc[i]:
                    hold = 0
                else: 
                    hold = 1
                rate_data.append(rate)
            rate_data = pd.Series(rate_data,index = dframe.index, name = 'st2')
            return rate_data
        def strategy_sma(dframe):
            rate = 100
            rate_data = [rate]  
            hold = 1
            for i in range(1,len(dframe['Close'])):
                if hold:
                    rate = rate*dframe['Close'].iloc[i]/dframe['Close'].iloc[i-1]
                if dframe['Close'].iloc[i] >= dframe['sma5'].iloc[i]:
                    hold = 1
                else: 
                    hold = 0
                rate_data.append(rate)
            rate_data = pd.Series(rate_data,index = dframe.index, name = 'st_sma5')
            return rate_data
        def strategy_macd(dframe):
            rate = 100
            rate_data = [rate]  
            hold = 1
            for i in range(1,len(dframe['Close'])):
                if hold:
                    rate = rate*dframe['Close'].iloc[i]/dframe['Close'].iloc[i-1]
                if dframe['macdhist'].iloc[i] >= 0:
                    hold = 1
                else: 
                    hold = 0
                rate_data.append(rate)
            rate_data = pd.Series(rate_data,index = dframe.index, name = 'st_macd')
            return rate_data
        def strategy_kd(dframe):
            rate = 100
            rate_data = [rate]  
            hold = 1
            for i in range(1,len(dframe['Close'])):
                if hold:
                    rate = rate*dframe['Close'].iloc[i]/dframe['Close'].iloc[i-1]
                if dframe['slowk'].iloc[i] >= dframe['slowd'].iloc[i]:
                    hold = 1
                else: 
                    hold = 0
                rate_data.append(rate)
            rate_data = pd.Series(rate_data,index = dframe.index, name = 'st_kd')
            return rate_data
        def strategy_rsi(dframe):
            rate = 100
            rate_data = [rate]  
            hold = 1
            for i in range(1,len(dframe['Close'])):
                if hold:
                    rate = rate*dframe['Close'].iloc[i]/dframe['Close'].iloc[i-1]
                if dframe['rsif'].iloc[i] >= dframe['rsis'].iloc[i]:
                    hold = 1
                else: 
                    hold = 0
                rate_data.append(rate)
            rate_data = pd.Series(rate_data,index = dframe.index, name = 'st_rsi')
            return rate_data

        totval = pd.concat([
            strategy0(res), 
            strategy1(res),
            strategy2(res),
            strategy_sma(res),
            strategy_macd(res),
            strategy_kd(res),
            strategy_rsi(res)
            ],axis = 1)

        draw_pic_totval(totval)
        


        res.to_excel(workbook_path, index=True) 
        workbook.close()
        return jsonify({'images': result, 'excel_path': workbook_path})
    except Exception as e:
        return jsonify({'error': f'處理股票數據時發生錯誤：{str(e)}'}), 500



class GPThandler:
    def __init__(self, image_paths, question):
        self.image_paths = image_paths
        self.question = question

    def extract_text_from_image(self):
        client = vision.ImageAnnotatorClient()
        extracted_texts = []
        for image_path in self.image_paths:
            with io.open(image_path, 'rb') as image_file:
                content = image_file.read()
            image = vision.Image(content=content)
            response = client.text_detection(image=image)
            texts = response.text_annotations
            if texts:
                extracted_texts.append(texts[0].description.strip())
            else:
                extracted_texts.append(f"No text found in the image: {image_path}")
        return '\n'.join(extracted_texts)

    def generate_gpt_response(self, combined_texts):
        prompt = f"Based on the following image data: {combined_texts}, please provide a detailed stock analysis after I ask the next question."
        gpt_response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a stock market analyst."},
                {"role": "system", "content": prompt},
                {"role": "user", "content": self.question},
            ],
            max_tokens=150,
            temperature=0.2
        )
        return gpt_response['choices'][0]['message']['content']

@app.route('/api/send_to_gpt', methods=['POST'])
def send_to_gpt():
    try:
        data = request.json
        image_paths = data.get('image_paths')
        question = data.get("question")

        if not image_paths or not question:
            return jsonify({'error': '缺少必需的參數'}), 400

        handler = GPThandler(image_paths, question)
        combined_texts = handler.extract_text_from_image()
        gpt_reply = handler.generate_gpt_response(combined_texts)

        return jsonify({'gpt_reply': gpt_reply})

    except Exception as e:
        return jsonify({'error': f'處理請求時發生錯誤：{str(e)}'}), 500


@app.route('/api/download_excel', methods=['GET'])
def download_excel():
    try:
        file_path = request.args.get('file_path')

        if not file_path or not os.path.exists(file_path):
            return jsonify({'error': '文件不存在'}), 400

        return send_file(file_path, as_attachment=True)
    except Exception as e:
        return jsonify({'error': f'下載Excel文件時發生錯誤：{str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)


