
from common.writeexcel import Write_excel
import json


def send_requests(s, testdata):
    # 封装requests请求
    method = testdata["method"]
    url = testdata["url"]
    # url后面的params参数
    try:
        params = eval(testdata["params"])
    except:
        params = None
    # 请求头部headers
    try:
        headers = eval(testdata["headers"])
        print("请求头部: %s" % headers)
    except:
        headers = None
    # post请求body类型
    type = testdata["type"]

    test_num = testdata['id']
    print("******正在执行用例: ------  %s  ------******" % test_num)
    print("请求方式: %s, 请求url: %s" % (method, url))
    print("请求params: %s" % params)

    # post请求body内容
    try:
        bodydata = eval(testdata["body"])
    except:
        bodydata = {}

    # 判断传data数据还是json
    if type == "data":
        body = bodydata
    elif type == "json":
        body = json.dumps(bodydata)
    else:
        body = bodydata
    if method == "post":
        print("post请求body类型为: %s, body内容为: %s" % (type, body))

    verify = False
    res = {}  # 接受返回数据

    try:
        r = s.request(method=method, 
                      url=url,
                      params=params,
                      headers=headers,
                      data=body,
                      verify=verify
                      )
        print("页面返回信息: %s" % r.content.decode("utf-8"))
        res['id'] = testdata['id']
        res['rowNum'] = testdata['rowNum']
        res["statuscode"] = str(r.status_code)  # 状态码转成str
        res["text"] = r.content.decode("utf-8")
        res["times"] = str(r.elapsed.total_seconds())  # 接口请求时间转换成str
        if res["statuscode"] != "200":
            res["error"] = res["text"]
        else:
            res["error"] = ""
        res["msg"] = ""
        if testdata["checkpoint"] in res["text"]:
            res["result"] = "pass"
            print("用例测试结果: %s---->%s" % (test_num, res["result"]))
        else:
            res["result"] = "fail"
        return res
    except Exception as msg:
        res["msg"] = str(msg)
        return res


def write_result(result, filename="result.xlsx"):
    # 返回结果的行数row_num
    row_num = result['rowNum']
    # 写入statuscode
    wt = Write_excel(filename)
    wt.write(row_num, 8, result['statuscode'])      # 写入返回状态码statuscode,第8列
    wt.write(row_num, 9, result['times'])          # 耗时
    wt.write(row_num, 10, result['error'])          # 状态码非200时的返回信息
    wt.write(row_num, 12, result['result'])         # 测试结果pass还是fail
    wt.write(row_num, 13, result['msg'])            # 抛异常
