
        这段是需要在百度智能云上获取你更改你自己的接口。每一个key都不一样 你要在百度智能云上创建

def get_access_token():
    """
    使用 AK，SK 生成鉴权签名（Access Token）
    :return: access_token，或是None(如果错误)
    """
    try:
        url = 'https://aip.baidubce.com/oauth/2.0/token'
        params = {"grant_type": "client_credentials", "client_id": API_KEY, "client_secret": SECRET_KEY}
        response = requests.post(url, params=params)
        if response.status_code == 200:
            access_token = response.json().get("access_token")
            return access_token
        else:
            print('获取 access_token 失败：', response.text)
    except Exception as e:
        print('获取 access_token 出错：', e)

        这段是需要在百度智能云上获取你更改你自己的接口。





 获得发票识别的类容
  def get_context(pic):
    try:
        # 二进制方式打开图片文件
        with open(pic, 'rb') as f:
            img = base64.b64encode(f.read())

        url = "https://aip.baidubce.com/rest/2.0/ocr/v1/vat_invoice"
        access_token = get_access_token()
        headers = {'content-type': 'application/x-www-form-urlencoded'}
        params = {"image":img, "access_token":access_token}
        response = requests.post(url, data=params, headers=headers)

        if response.status_code == 200:
            json1 = response.json()
            if 'words_result' in json1:
                data = {}
                data['发票代码']=json1['words_result'].get( "InvoiceNumConfirm")
                data["购买方纳税人识别号"] = json1['words_result'].get("PurchaserRegisterNum", '')
                data['购买方名称'] = json1['words_result'].get('PurchaserName', '')
                data['销售方名称'] = json1['words_result'].get('SellerName', '')
                data['销售方纳税人识别号'] = json1['words_result'].get('SellerRegisterNum', '')
                data['金额(不含税)'] = json1['words_result'].get('TotalAmount', '') #金额
                data["税率"] = str(json1['words_result'].get('CommodityTaxRate'))
                data['税额']=json1['words_result'].get('TotalTax','')
                data['价税合计(含税价格)']=json1['words_result'].get('AmountInFiguers')
                data["项目名称"] = str(json1['words_result'].get("CommodityName", ''))









                return data
            else:
                print('无法识别发票内容：', pic)
        else:
            print('请求识别发票内容失败：', pic)
    except Exception as e:
        print('识别发票内容出错：', pic, e)




  将你需要的识别信息导入在excel表里面
     def save_to_excel(datas):
    print('正在写入数据！')
    book = openpyxl.Workbook()
    sheet = book.active
    title = ['购买方名称', '销售方名称', '购买方纳税人识别号', '销售方纳税人识别号', '金额(不含税)','税额','价税合计(含税价格)','发票代码','项目名称']
    sheet.append(title)
    for data in datas:
        row = [data.get(field, '') for field in title]
        sheet.append(row)
    book.save('增值税发票.xlsx')
    print('数据写入成功！')、
  将你需要的识别信息导入在excel表里面
