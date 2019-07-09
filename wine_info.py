"""
Authors: Wang Fei
History: 2019.07.08
Dependencies:
    requests
    beautifulsoup4
        lxml
        html5lib
    pandas
        xlsxwriter
        xlrd
"""
import os
import time
import re
import bs4
import requests
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime


def wineyun_extract(hostname="http://www.wineyun.com/",
                    save_path="wineyun.xlsx", img_dir="./imgs"):
    parse_attrs = ["品名", "产区", "品种", "类型", "容量"]   # from the table in the wine section
    if not os.path.exists(img_dir):
        os.makedirs(img_dir)

    check_update = False
    have_update = False
    orig_titles = list()
    if os.path.exists(save_path):
        # orig_df = pd.read_csv(save_path)
        orig_df = pd.read_excel(save_path)
        orig_titles = orig_df["标题"]
        dump_df = orig_df
        check_update = True
    else:
        dump_df = pd.DataFrame(columns=["标题", ] + parse_attrs +
                                       ["酒庄", "价格", "产品链接", "图片", "更新日期"])

    form_pages = range(240, -1, -40)
    for page in form_pages:
        if page == 0:
            response = requests.get(hostname)
        else:
            response = requests.post(hostname, data={"page": page})
        html = response.text
        soup = BeautifulSoup(html, "html5lib")

        # paths of each wine for sale.
        wine_path_list = list()

        # check if need to look further
        if check_update:
            title_tags = soup.find_all('h1', attrs={'class': 'bti ml10'})
            if len(title_tags) > 0:
                for title_tag in title_tags:
                    wine_title = title_tag.text
                    if wine_title not in orig_titles.values:
                        # it seems stupid because next_sibling can read white spaces between tags.
                        for sibling in title_tag.next_siblings:
                            if type(sibling) is bs4.element.Tag:
                                wine_path_list.append(sibling.dt.a["href"])
                                break
        else:
            # don't check, just download the information in all pages.
            dt_list = soup.find_all('dt', attrs={'class': 'fl'})
            for dt in dt_list:
                # rmb_price = dt.find("i", attrs={'class': "fl"}).text
                wine_path_list.append(dt.a["href"])

        if len(wine_path_list) == 0:
            continue
        else:
            have_update = True
        wine_url_list = ["%s/%s" % (hostname, p) for p in wine_path_list]

        for wine_url in wine_url_list:
            goods_id = wine_url.split(r"/")[-1]
            child_response = requests.get(wine_url)
            child_soup = BeautifulSoup(child_response.text, "html5lib")
            info_dict = dict()

            # # step 1: get the title; the title seems always exists.
            title = child_soup.head.title.text
            info_dict.update({"标题": title})
            print(30*"*" + title + 30*"*")
            # print("url: {}".format(wine_url))

            # # step 2: get the wine section
            trouble_attrs = list()
            for attr in parse_attrs:
                td1 = child_soup.find("td", text=attr)
                if td1 is None:
                    # sometimes the page use "种类" instead of "品种"
                    if attr == "品种":
                        td1 = child_soup.find("td", text="种类")
                        if td1 is None:
                            trouble_attrs.append(attr)
                            info_dict.update({attr: ""})
                        else:
                            info_dict.update({attr: td1.next_sibling.text})
                    else:
                        trouble_attrs.append(attr)
                        info_dict.update({attr: ""})
                else:
                    info_dict.update({attr: td1.next_sibling.text})

            # Maybe it is a bundle of wines
            if len(trouble_attrs) > 4:
                wine_section = child_soup.find("section", id="wine")
                if wine_section is not None:
                    # assmue the name is wrapped with the "span" tag,
                    wine_name_tags = wine_section("span")
                    if len(wine_name_tags) > 0:
                        wine_names = [wine_name_tags[i].text.replace("\n", " ").strip()
                                      for i in range(len(wine_name_tags))]
                        info_dict.update({"品名": " && ".join(wine_names)})
                        if "品名" in trouble_attrs:
                            trouble_attrs.remove("品名")
                    else:
                        # otherwise assume the first "p" tag contains the name.
                        wine_section_paragraphs = wine_section("p")
                        if len(wine_section_paragraphs) > 0:
                            wine_name = wine_section("p")[0].text.replace("\n", " ").strip()
                            info_dict.update({"品名": wine_name})
                            if "品名" in trouble_attrs:
                                trouble_attrs.remove("品名")
                        else:
                            # sometimes the wine section is a whole <div> block, not separated by <p>
                            # tags, such as http://www.wineyun.com///group/44446.
                            # No handling for this situation yet.
                            pass
                    # get other attributes with regular expression
                    for attr in parse_attrs[1:]:
                        attr_values = re.findall('{}：(.*?)\n'.format(attr), wine_section.text)
                        if len(attr_values) > 0:
                            info_dict.update({attr: " && ".join(attr_values)})
                            if attr in trouble_attrs:
                                trouble_attrs.remove(attr)

            # # step 3: save the url and get the picture of the wine
            info_dict.update({"产品链接": wine_url})
            pic_url = child_soup.find("img", id="showimgurl0")["src"]
            # TODO: download the pic
            pic_response = requests.get(pic_url)
            local_pic_path = os.path.join(img_dir, "{}.png".format(goods_id))
            if not os.path.exists(local_pic_path):
                with open(local_pic_path, "wb") as fin:
                    fin.write(pic_response.content)
            # must use double quotations and relative path(not absolute path in Windows)
            info_dict.update({"图片": '=HYPERLINK("{0}", "{1}")'.format(local_pic_path,
                                                                      "{}.png".format(goods_id))})
            # info_dict.update({"图片": pic_url})

            # # step 4: get the winery section
            winery_section = child_soup.find("section", id="winery")
            if winery_section is None:
                trouble_attrs.append("酒庄")
            else:
                winery = winery_section.find("a", target="_blank").text
                info_dict.update({"酒庄": winery})

            # # step 5: get price using regular expression
            search_obj = re.search('unitprice=\"(.*?)\"', child_soup.text)
            price = float(search_obj.group(1))
            info_dict.update({"价格": price})

            # # step 6: add the datetime
            info_dict.update({"更新日期": datetime.now().strftime("%x %X")})

            # # step 7: print logs
            if len(trouble_attrs) > 0:
                print("Error in getting (%s): %s" % (" ".join(trouble_attrs), title))
                print("url: {}".format(wine_url))

            # # step 8: save to pandas.DataFrame
            dump_df = dump_df.append(info_dict, ignore_index=True)
            time.sleep(1)
            print(info_dict)

    # TODO: sort and save
    # print(dump_df)
    if have_update:
        # dump_df.to_csv(save_path, index=False, encoding="GB18030")
        # write to excel(.xlsx)
        writer = pd.ExcelWriter(save_path, engine='xlsxwriter')
        dump_df.sort_values(by="更新日期", ascending=False).to_excel(writer, sheet_name='Sheet1', index=False, )
        # worksheet = writer.sheets['Sheet1']
        # worksheet.insert_image("J1", "test.png")
        writer.save()


if __name__ == "__main__":
    wineyun_extract()
