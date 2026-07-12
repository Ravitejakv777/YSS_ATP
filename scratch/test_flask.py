from flask import Flask
app = Flask('test')

@app.route('/')
def hello():
    return 'ok'

if __name__ == '__main__':
    try:
        print("Starting test server...")
        app.run(port=8082)
    except Exception as e:
        print("CATCHED EXCEPTION:")
        import traceback
        traceback.print_exc()
