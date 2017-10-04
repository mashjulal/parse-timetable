import xlwt


class Styles:

    TITLE = xlwt.easyxf("font: bold on; "
                        "borders: left medium, right medium, top medium, bottom medium;"
                        "alignment: horiz center")

    NON_STUDY = xlwt.easyxf("pattern: pattern solid, fore_color blue;"
                            "font: bold on;"
                            "borders: left medium, right medium, top medium, bottom medium;"
                            "alignment: horiz center")

    WEEKDAY = xlwt.easyxf("font: bold on;"
                          "borders: left medium, right medium, top medium, bottom medium;"
                          "alignment: horiz center, vert center, rota 90")

    PRACTICE = xlwt.easyxf("pattern: pattern solid, fore_color green;"
                           "font: bold on;"
                           "borders: left medium, right medium, top medium, bottom medium;"
                           "alignment: horiz center")

    LECTURE = xlwt.easyxf("pattern: pattern solid, fore_color yellow;"
                          "font: bold on;"
                          "borders: left medium, right medium, top medium, bottom medium;"
                          "alignment: horiz center")

    LAB = xlwt.easyxf("pattern: pattern solid, fore_color blue;"
                      "font: bold on;"
                      "borders: left medium, right medium, top medium, bottom medium;"
                      "alignment: horiz center")
