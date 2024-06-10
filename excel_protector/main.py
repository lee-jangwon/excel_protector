import argparse
from spire.xls import *
from spire.xls.common import *
from pathlib import Path
from password_generator import PasswordGenerator

def read_workbook(filename: Path) -> Workbook:
    workbook = Workbook()
    workbook.LoadFromFile(str(filename))
    print(workbook)
    ## 비밀번호 생성
    password_machine = PasswordGenerator()
    password_machine.minlen = 15
    password_machine.maxlen = 15

    password = password_machine.generate()
    print(f"비밀번호는 한 번밖에 보이지 않고 저장되지 않습니다. 꼭 저장해주세요!\n비밀번호:\t{password}")
    workbook.Protect(password)

    save_path = filename.parent / f"{filename.stem}_protected.xlsx"
    workbook.SaveToFile(str(save_path), ExcelVersion.Version2016)
    workbook.Dispose()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="뉴스레터에 보낼 파일을 암호화합니다.")
    parser.add_argument("filename", type=str, help="엑셀파일이 있는 절대 경로를 넣어주세요. 만약 공백이 있다면 따옴표를 넣어주세요.\nExcel Path:\t")

    args = parser.parse_args()
    filepath = Path(args.filename)

    if not filepath.exists() or not filepath.is_file():
        print(f"에러: {filename}이 없네용.")

    read_workbook(filepath)