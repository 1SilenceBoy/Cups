import requests


def create_get_printer_ipp(ipp_printer_url,server_host):
    # version : 1.0
    # Get-Printer-Attributes
    # 1
    # attributes-charset (charset): 'utf-8'
    # attributes-natural-language (naturalLanguage): 'en-us'
    request_body = (b'\x01\x00\x00\x0b\x00\x00\x00\x0b\x01\x47\x00\x12\x61\x74\x74\x72\x69\x62\x75\x74\x65\x73\x2d'
                    b'\x63\x68\x61\x72\x73\x65\x74\x00\x05\x75\x74\x66\x2d\x38\x48\x00\x1b\x61\x74\x74\x72\x69'
                    b'\x62\x75\x74\x65\x73\x2d\x6e\x61\x74\x75\x72\x61\x6c\x2d\x6c\x61\x6e\x67\x75\x61\x67\x65'
                    b'\x00\x05\x65\x6e\x2d\x75\x73')
    # printer-uri
    request_body += b'\x45\x00\x0b\x70\x72\x69\x6e\x74\x65\x72\x2d\x75\x72\x69\x00'
    request_body += bytes([len(ipp_printer_url)])
    request_body += ipp_printer_url

    # end
    request_body += b'\x03'

    # 发送IPP请求
    response_code = requests_post(ipp_printer_url, request_body, server_host)
    return response_code


def print_file(ipp_printer_url, file_path, job_name, user_name, server_host):
    # version : 1.0
    # Print-Job
    # 1
    # attributes-charset (charset): 'utf-8'
    # attributes-natural-language (naturalLanguage): 'en-us'
    request_body = (b'\x01\x00\x00\x02\x00\x00\x00\x02\x01\x47\x00\x12\x61\x74\x74\x72\x69\x62\x75\x74\x65\x73\x2d'
                    b'\x63\x68\x61\x72\x73\x65\x74\x00\x05\x75\x74\x66\x2d\x38\x48\x00\x1b\x61\x74\x74\x72\x69'
                    b'\x62\x75\x74\x65\x73\x2d\x6e\x61\x74\x75\x72\x61\x6c\x2d\x6c\x61\x6e\x67\x75\x61\x67\x65'
                    b'\x00\x05\x65\x6e\x2d\x75\x73')
    # printer-uri
    request_body += b'\x45\x00\x0b\x70\x72\x69\x6e\x74\x65\x72\x2d\x75\x72\x69\x00'
    request_body += bytes([len(ipp_printer_url)])
    request_body += ipp_printer_url

    # job name
    request_body += b'\x42\x00\x08\x6a\x6f\x62\x2d\x6e\x61\x6d\x65\x00'
    request_body += bytes([len(job_name)]) + job_name

    # requesting-user-name
    request_body += (b'\x42\x00\x14\x72\x65\x71\x75\x65\x73\x74\x69\x6e'
                     b'\x67\x2d\x75\x73\x65\x72\x2d\x6e\x61\x6d\x65\x00')
    request_body += bytes([len(user_name)]) + user_name

    # end
    request_body += b'\x02\x03'

    with open(file_path, 'rb') as f:
        data = f.read()
    # 数据体构成
    request_body += data

    # 发送IPP请求
    response = requests_post(ipp_printer_url, request_body, server_host)
    return response


def requests_post(ipp_printer_url, ipp_request_body, server_host):
    # IPP请求头部
    headers = {
        'Content-Type': 'application/ipp',
        'User-Agent': 'Internet Print Provider',
        'Connection': 'Keep-Alive',
        'Content-Length': str(len(ipp_request_body)),
        'Host': server_host+':631'
    }
    response = requests.post(ipp_printer_url, headers=headers, data=ipp_request_body)

    # 检查响应状态码并打印响应内容
    # if response.status_code == 200:
    #     print("IPP请求成功")
    #     # print(response.text)  # 打印响应的文本内容（注意：这通常是XML格式）
    # else:
    #     print(f"IPP请求失败，状态码：{response.status_code}")
    #     print(response.text)  # 打印任何可能的错误消息
    return response.status_code
