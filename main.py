from flask import Flask, jsonify, request

from excel_class import ExcelMain

app = Flask(__name__)


@app.route('/<r:int>',methods=["POST"])
def hello_world(r):
    print("开始")
    params = request.json
    print(params)
    ex=ExcelMain(params)
    dict=ex.get_return()
    print("结束")
    return jsonify(code=200, message='ok',data=dict)


if __name__ == '__main__':
    import sys,threading
    sys.setrecursionlimit(100000)
    threading.stack_size(200000000)
    # thread = threading.Thread(target=app.run,args=("0.0.0.0",8055))
    # thread.start()
    # print(app.url_map)
    app.run("0.0.0.0",8055)

