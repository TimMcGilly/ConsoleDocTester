import csv
import inspect
import os
import unittest
from enum import Enum
from typing import List
from pathlib import Path

import pyperclip
from pywinauto.application import Application
import glob


class TestType(Enum):
    Typical = 1
    Erroneous = 2
    Extreme = 3


globals().update(TestType.__members__)


class ConsoleTester:
    def __init__(self, appPath, testOutDir, title=None):
        self.testCase = []
        self.appPath = appPath
        self.testOutDir = testOutDir
        if title is None:
            self.title = appPath
        else:
            self.title = title

    def findNextImgPath(self):
        files = glob.glob(self.testOutDir + "/*.png")
        maxFileNum = 0
        for file in files:
            val = int(Path(file).stem[3:])
            if val > maxFileNum:
                maxFileNum = val

        return self.testOutDir + f"/Img{maxFileNum + 1}.png"

    def UpdateTestCase(self, funcName, description: str, testData: str, testType: TestType, expectedValue: str,
                       actualValue: str, testPass: bool, imgPath: str):
        testPassStr = "Pass" if testPass else "Fail"

        outputFilePath = Path(self.testOutDir + '/testResults.csv',)

        testRow = None

        maxTestNum = 1

        if outputFilePath.is_file():
            with open(outputFilePath, 'r', newline='') as csvfile:
                csv_reader = csv.reader(csvfile, delimiter=' ', quotechar='|')
                rows = list(csv_reader)

                index = 0
                for row in rows:
                    if int(row[1]) > maxTestNum:
                        maxTestNum = int(row[1])
                    if row[0] == funcName:
                        testRow = row
                        index += 1

        if testRow is None:
            with open(outputFilePath, 'a', newline='') as csvfile:
                csv_writer = csv.writer(csvfile)
                csv_writer.writerow(
                    [funcName, str(maxTestNum + 1), description, testData,
                     testType.name, expectedValue, actualValue, testPassStr, imgPath, False, None])
        else:
            with open(outputFilePath, 'w', newline='') as csvfile:
                if testRow[7] == "Fail" and testPassStr == "Pass":
                    # HasCorrection
                    rows[index][9] = True
                    # ImgPath2
                    rows[index][10] = imgPath
                else:
                    rows[index][8] = imgPath

                rows[index][2] = description
                rows[index][3] = testData
                rows[index][4] = testType.name
                rows[index][5] = expectedValue
                rows[index][6] = actualValue
                rows[index][7] = testPassStr

                csvfile.seek(0)
                csv_writer = csv.writer(csvfile)
                csv_writer.writerows(rows)

    def addTestCase(self, testType: TestType, description: str, testData: List[str], expectedValue: List[str]):
        app = Application().start(self.appPath, create_new_console=True, wait_for_idle=False)

        for data in testData:
            app.window().type_keys(data + "~", pause=0.1
                                   )

        imgPath = self.findNextImgPath()
        app.window().capture_as_image().save(imgPath)
        app.window().type_keys(r'^a')
        app.window().type_keys(r'^c')
        app.kill()

        lines = pyperclip.paste().splitlines()
        index = lines.index(testData[-1])
        actual = lines[index + 1: index + 1 + len(expectedValue)]

        self.UpdateTestCase(inspect.currentframe().f_back.f_code.co_name, description, " ".join(testData), testType, " ".join(expectedValue), " ".join(actual), actual==expectedValue, imgPath)
        return actual, expectedValue


class TestPasswordChecker(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tester = ConsoleTester(
            "C:\\Users\\timmc\\source\\repos\\PasswordChecker\\PasswordChecker\\bin\\Debug\\netcoreapp3.1\\PasswordChecker.exe",
            "C:\\Users\\timmc\\source\\repos\\PasswordChecker\\PasswordChecker\\testOutput")

    def test_TooShortPasswordValidation(self):
        testType = Erroneous
        description = "Password must be greater than 8 character"
        testData = ["1", "aaa"]
        expectedValue = ["Password length is not between 8 and 24 inclusive."]

        self.assertEqual(*self.tester.addTestCase(testType, description, testData, expectedValue))

    def test_isupper(self):
        self.assertTrue('FOO'.isupper())
        self.assertFalse('Foo'.isupper())

    def test_split(self):
        s = 'hello world'
        self.assertEqual(s.split(), ['hello', 'world'])
        # check that s.split fails when the separator is not a string
        with self.assertRaises(TypeError):
            s.split(2)

    @classmethod
    def tearDownClass(cls):
        print("Destroy")


if __name__ == '__main__':
    unittest.main()
