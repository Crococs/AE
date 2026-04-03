import argparse

parser = argparse.ArgumentParser(description="인사하는 프로그램")

parser.add_argument('--name', type=str, default="Stranger", help="이름을 입력하세요")
parser.add_argument('--age', type=int, default=30, help="나이를 입력하세요")

args = parser.parse_args()

print(f"Hello, {args.name} {args.age}! AI 엔지니어의 세계에 오신 것을 환영합니다.")