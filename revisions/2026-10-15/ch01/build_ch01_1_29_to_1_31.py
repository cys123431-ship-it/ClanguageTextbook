from pathlib import Path
import json, subprocess, zipfile, hashlib, datetime, re
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT=Path(__file__).parent
SRC=ROOT/'sources'; SRC.mkdir(exist_ok=True)
FLAGS=['-std=c11','-Wall','-Wextra','-Wpedantic','-Werror','-O0']

# Every example is authored independently. Output is checked with both GCC and Clang.
sections=[
 dict(num='1.29',title='전처리기는 언제 일하는가: #include와 이름 붙은 상수',
 intro=[
 'C 프로그램을 컴파일할 때 컴파일러가 곧바로 모든 문장을 실행 파일로 바꾸는 것은 아니다. 소스의 전처리 지시문을 먼저 처리하여 컴파일 단계에 전달할 번역 단위의 내용을 준비한다. #으로 시작하는 지시문은 C 실행문이 아니라 번역 전의 지시다. 따라서 실행 중 if처럼 조건을 검사하거나 변수를 만드는 명령으로 오해하면 안 된다.',
 '헤더는 여러 소스 파일이 함께 알아야 하는 선언, 자료형, 매크로를 제공한다. #include <stdio.h>는 표준 입출력 인터페이스를 사용할 수 있게 관련 헤더를 포함시키도록 지시한다. 실제 탐색 경로와 물리적 포함 방식은 구현의 규칙에 의존하므로, 무조건 같은 이름의 파일을 한 폴더에서 읽는다고 단정할 수 없다.',
 '매크로는 변수와 다르다. #define LIMIT 5에서 LIMIT은 저장 공간을 가진 int 객체가 아니다. 전처리 시점에 대응하는 전처리 토큰으로 바뀐다. 숫자에 이름을 붙여 의미를 밝히는 용도는 유용하지만, 타입 검사나 주소가 필요한 객체를 대체할 수는 없다.',
 'C11의 번역 단계는 개념적으로 여러 단계로 설명된다. 주석은 공백으로 대체되고 전처리 지시문과 매크로 확장 등이 처리된 다음, C 문법 분석과 링크를 거쳐 프로그램이 구성된다. 전처리 결과를 확인하는 gcc -E, clang -E는 실습용 컴파일러 옵션이지 ISO C의 문법이 아니다.'
 ], terms=[('전처리 지시문','실행 중 수행되는 명령이 아니라 컴파일 이전 번역 과정에 영향을 주는 문장이다.'),('객체형 매크로','인수를 받지 않고 이름이 대응하는 전처리 토큰열로 확장되는 매크로다.'),('번역 단위','전처리 결과가 반영되어 컴파일 대상이 되는 소스 단위다.'),('헤더','함수 원형·매크로·자료형 같은 인터페이스를 여러 파일에 알리는 매개다.')],
 examples=[
 dict(id='29a',level='완성 예제',title='표준 헤더가 제공하는 printf 선언과 출력',why='표준 함수의 선언 없이 사용하면 C11에서는 함수 호출을 올바르게 검증할 수 없다. 따라서 표준 헤더를 먼저 포함하는 습관이 중요하다.',files={'main.c':'''#include <stdio.h>\n\nint main(void)\n{\n    printf("include works\\n");\n    return 0;\n}\n'''},expected='include works\n',predict='첫 줄은 화면에 출력되는가? 실제로 출력하는 문장은 어디인가?',explain=['#include <stdio.h>는 stdio.h 헤더의 인터페이스를 사용할 수 있도록 전처리한다. 이 줄 자체는 출력을 발생시키지 않는다.','int main(void)는 프로그램 진입 함수의 정의를 시작한다. void는 인수를 받지 않음을 명시한다.','printf는 문자열을 표준 출력에 기록한다. 문자열 속 \\n은 줄바꿈 문자다.','return 0은 정상 종료 상태를 나타낸다.'],trace='전처리 단계: 헤더 관련 내용 준비 → 컴파일: printf 선언을 토대로 호출 검사 → 실행: printf 한 번 호출 → 종료 상태 0.',mistake='헤더를 포함하는 것과 라이브러리 함수를 호출하는 것은 다르다. #include만 적으면 아무 글자도 출력되지 않는다.',variant='출력 문자열을 header checked로 바꾸어 보자.',variant_files={'main.c':'''#include <stdio.h>\nint main(void)\n{\n    printf("header checked\\n");\n    return 0;\n}\n'''},variant_expected='header checked\n',question='stdio.h는 파일 내용을 출력하는 함수인가?',answer='아니다. 헤더는 선언 등의 인터페이스를 제공한다. 출력은 printf 호출이 담당한다.'),
 dict(id='29b',level='완성 예제',title='상수의 이름을 바꾸면 반복 횟수도 바뀐다',why='반복 상한이 여러 곳에 숫자로 흩어지면 유지보수가 어려워진다. 상수에 의미 있는 이름을 붙여 변경 지점을 줄인다.',files={'main.c':'''#include <stdio.h>\n#define DAYS 4\n\nint main(void)\n{\n    int total = 0;\n    for (int day = 1; day <= DAYS; ++day) {\n        total += day;\n        printf("day=%d total=%d\\n", day, total);\n    }\n    return 0;\n}\n'''},expected='day=1 total=1\nday=2 total=3\nday=3 total=6\nday=4 total=10\n',predict='DAYS는 실제 int 변수인가? 반복은 몇 번 일어나고 최종 total은 얼마인가?',explain=['#define DAYS 4는 이후의 DAYS 전처리 토큰을 4로 확장한다. 세미콜론을 붙이지 않는다.','int total = 0은 합계 객체를 초기화한다.','for의 day는 1에서 시작해 DAYS 이하일 동안 반복하고 한 번마다 1씩 증가한다.','total += day는 누적 합계를 갱신하며 printf는 갱신된 값을 확인하는 학습용 출력이다.'],trace='day=1: 0+1=1 → day=2: 1+2=3 → day=3: 3+3=6 → day=4: 6+4=10 → day=5에서 조건 거짓.',mistake='#define DAYS 4;처럼 세미콜론을 넣으면 DAYS를 쓰는 자리에 4; 토큰이 들어가 식의 문법이 깨질 수 있다.',variant='DAYS를 3으로 바꾸고 예상 출력을 직접 적은 뒤 실행하자.',variant_files={'main.c':'''#include <stdio.h>\n#define DAYS 3\nint main(void)\n{\n    int total = 0;\n    for (int day = 1; day <= DAYS; ++day) {\n        total += day;\n        printf("day=%d total=%d\\n", day, total);\n    }\n    return 0;\n}\n'''},variant_expected='day=1 total=1\nday=2 total=3\nday=3 total=6\n',question='DAYS에 &DAYS를 적용해 주소를 얻을 수 있는가?',answer='아니다. DAYS는 객체가 아니라 숫자 토큰으로 확장되므로 &4와 같은 형태가 되어 적절하지 않다.'),
 dict(id='29c',level='부분 완성 예제',title='문자열 매크로와 화면에 보이는 문구의 관계',why='프로그램에서 반복되는 메시지를 하나의 이름으로 관리할 수 있다. 다만 매크로는 자료형이 없는 토큰 확장이라는 점을 이해해야 한다.',files={'main.c':'''#include <stdio.h>\n#define LABEL "status"\n\nint main(void)\n{\n    int count = 2;\n    printf("%s=%d\\n", LABEL, count);\n    return 0;\n}\n'''},expected='status=2\n',predict='LABEL은 실행 중 문자열을 조립하는 함수인가? printf의 %s는 무엇을 받는가?',explain=['LABEL은 "status"라는 문자열 리터럴 토큰으로 확장된다.','count는 실제 저장 공간을 가진 int 지역 객체다.','printf의 %s는 문자열을 출력하고 %d는 int 값을 출력한다. 형식 지정자와 인수의 자료형을 맞춰야 한다.'],trace='전처리 후 printf 호출은 문자열 리터럴 "status"와 int 객체 count의 현재 값 2를 인수로 사용한다. 출력 후 count는 바뀌지 않는다.',mistake='문자열 리터럴 매크로를 int 변수와 같은 저장 객체로 취급하지 않는다.',variant='LABEL을 "items"로 바꾸고 count를 5로 초기화하자.',variant_files={'main.c':'''#include <stdio.h>\n#define LABEL "items"\nint main(void)\n{\n    int count = 5;\n    printf("%s=%d\\n", LABEL, count);\n    return 0;\n}\n'''},variant_expected='items=5\n',question='LABEL의 타입을 #define 지시문 자체가 선언하는가?',answer='아니다. 매크로는 타입을 선언하지 않는다. 확장된 문자열 리터럴은 C 언어의 타입 규칙에 따라 해석된다.'),
 dict(id='29d',level='독립 작성 예제',title='두 개의 상수를 조합해 배송비 계산',why='프로그램 정책을 상수로 관리하고 계산은 지역 객체에서 수행한다. 숫자 정책과 계산 로직을 분리하면 검토하기 쉽다.',files={'main.c':'''#include <stdio.h>\n#define UNIT_PRICE 1200\n#define COUNT 3\n\nint main(void)\n{\n    int cost = UNIT_PRICE * COUNT;\n    printf("cost=%d\\n", cost);\n    return 0;\n}\n'''},expected='cost=3600\n',predict='UNIT_PRICE와 COUNT 중 실제 저장 객체가 있는 것은 무엇인가? cost는 얼마인가?',explain=['UNIT_PRICE와 COUNT는 정수 리터럴로 확장되는 이름이다.','int cost = UNIT_PRICE * COUNT;는 확장된 1200 * 3을 계산해 지역 객체 cost를 초기화한다.','printf는 정책 계산 결과를 확인하기 위한 출력이며 계산의 본질은 곱셈과 초기화다.'],trace='전처리: UNIT_PRICE→1200, COUNT→3. 실행: 1200×3=3600을 cost에 저장하고 한 번 출력한다.',mistake='실제 업무에서 비용 곱셈의 정수 범위를 검사해야 한다. 여기서는 작은 상수라 int 범위를 넘지 않는다.',variant='UNIT_PRICE를 900으로, COUNT를 4로 바꾼 완성 프로그램을 작성하자.',variant_files={'main.c':'''#include <stdio.h>\n#define UNIT_PRICE 900\n#define COUNT 4\nint main(void)\n{\n    int cost = UNIT_PRICE * COUNT;\n    printf("cost=%d\\n", cost);\n    return 0;\n}\n'''},variant_expected='cost=3600\n',question='상수의 이름을 바꾸는 것과 실행 중 변수의 값을 바꾸는 것은 같은가?',answer='아니다. 전자는 번역 시점의 토큰 대체이고 후자는 실행 시점의 객체 값 변경이다.')
 ]),
 dict(num='1.30',title='함수형 매크로: 괄호, 인수 평가, 부작용',
 intro=[
 '함수형 매크로는 이름 뒤의 괄호에 전달한 인수를 매크로 본문에 대응시켜 확장한다. 이름과 여는 괄호 사이에 공백을 두면 객체형 매크로로 해석되므로 정의 문법이 중요하다. 매크로 호출의 겉모습은 함수 호출과 비슷해도 자료형 검사, 인수 평가 횟수, 디버깅 특성은 다르다.',
 '매크로 본문에서 인수마다 괄호를 감싸고 전체 식에도 괄호를 두는 이유는, 매크로 확장 뒤에 C 연산자 우선순위가 그대로 적용되기 때문이다. 이 원칙을 지키지 않으면 호출 위치의 식이 예상과 다른 묶음으로 해석될 수 있다.',
 '함수형 매크로의 가장 중요한 위험은 인수가 한 번만 평가된다는 보장이 없다는 점이다. 인수에 ++, --, 대입 또는 함수 호출을 넣으면 확장된 위치마다 평가가 발생할 수 있다. 특히 수정과 읽기가 순서 관계 없이 충돌하면 정의되지 않은 동작도 생길 수 있다. 이 절의 실습은 정의된 동작만 사용하여 차이를 관찰한다.',
 '간단한 연산이라도 타입 안전성과 단일 평가가 필요하면 함수가 더 적합하다. 아래 예제는 매크로를 무조건 권장하기 위한 것이 아니라, 기존 코드의 매크로를 정확히 읽고 안전하게 바꾸기 위한 것이다.'
 ],terms=[('함수형 매크로','인수를 받아 전처리 단계에서 대응 토큰열로 확장되는 매크로다.'),('연산자 우선순위','확장된 식에서 어떤 연산이 먼저 묶이는지 결정하는 문법 규칙이다.'),('부작용','표현식 평가 과정에서 객체 값 변경 등 관찰 가능한 상태 변화를 일으키는 효과다.'),('다중 평가','하나의 인수 표현식이 확장 결과에서 두 번 이상 평가되는 현상이다.')],
 examples=[
 dict(id='30a',level='완성 예제',title='괄호가 빠진 제곱 매크로의 뜻밖의 결과',why='매크로 본문을 문자처럼 끼워 넣는 방식으로 생각하면 연산 우선순위 오류를 미리 찾아낼 수 있다.',files={'main.c':'''#include <stdio.h>\n#define SQUARE_BAD(x) x * x\n#define SQUARE_GOOD(x) ((x) * (x))\n\nint main(void)\n{\n    int bad = SQUARE_BAD(2 + 3);\n    int good = SQUARE_GOOD(2 + 3);\n    printf("bad=%d good=%d\\n", bad, good);\n    return 0;\n}\n'''},expected='bad=11 good=25\n',predict='SQUARE_BAD(2+3)이 25가 아닌 이유를 확장된 식으로 먼저 적어 보자.',explain=['SQUARE_BAD(x)의 본문에는 인수 괄호가 없다. SQUARE_BAD(2 + 3)은 2 + 3 * 2 + 3으로 확장되어 11이 된다.','SQUARE_GOOD(x)는 인수와 전체를 괄호로 감싸 ((2 + 3) * (2 + 3))을 만든다.','두 매크로는 모두 전처리 시점의 토큰 확장이지 함수 호출이 아니다.','printf는 두 결과를 한 줄에 나란히 보여 오개념을 진단하는 학습용 출력이다.'],trace='bad: 곱셈 3×2=6 → 덧셈 2+6+3=11. good: 각각 2+3=5 → 5×5=25.',mistake='매크로 이름에 SQUARE가 있다고 해서 C가 제곱 함수로 취급하는 것은 아니다. 실제 확장식을 확인해야 한다.',variant='인수를 1 + 4로 바꾸어 잘못된 결과와 올바른 결과를 비교하자.',variant_files={'main.c':'''#include <stdio.h>\n#define SQUARE_BAD(x) x * x\n#define SQUARE_GOOD(x) ((x) * (x))\nint main(void)\n{\n    printf("bad=%d good=%d\\n", SQUARE_BAD(1 + 4), SQUARE_GOOD(1 + 4));\n    return 0;\n}\n'''},variant_expected='bad=9 good=25\n',question='인수에 괄호만 넣으면 항상 안전한가?',answer='아니다. 다중 평가와 자료형 문제가 여전히 남는다. 괄호는 우선순위 문제를 줄일 뿐이다.'),
 dict(id='30b',level='완성 예제',title='최댓값 매크로와 증가 연산의 다중 평가',why='매크로에 부작용 있는 인수를 전달할 때 발생하는 실제 위험을 관찰한다. 정의된 평가 순서를 가진 조건 연산자를 사용한다.',files={'main.c':'''#include <stdio.h>\n#define MAX_MACRO(a, b) ((a) > (b) ? (a) : (b))\n\nint main(void)\n{\n    int count = 4;\n    int result = MAX_MACRO(++count, 3);\n    printf("count=%d result=%d\\n", count, result);\n    return 0;\n}\n'''},expected='count=6 result=6\n',predict='++count는 한 번 실행될까, 두 번 실행될까? 조건 연산자의 선택 규칙을 생각하자.',explain=['MAX_MACRO(a,b)는 조건식에서 a를 한 번, 참 분기에서 a를 다시 사용한다.','count=4에서 ++count는 5가 되고 5>3은 참이다. 조건 연산자는 조건식을 먼저 평가하고 선택된 분기를 평가한다.','참 분기의 ++count가 다시 수행되어 6이 되고 결과에 저장된다.','이 특정 코드는 조건식 평가와 선택 분기 평가 사이에 순서가 있어 정의된 동작이지만, 매크로에 부작용 있는 인수를 넣는 일반 습관은 피해야 한다.'],trace='초기 count=4 → 조건에서 ++count: 5 → 5>3 참 → 참 분기에서 ++count: 6 → result=6.',mistake='매크로가 함수처럼 인수를 한 번 계산해 매개변수에 저장한다고 생각하면 count=5라고 잘못 예측한다.',variant='매크로 호출의 인수를 ++count가 아닌 count로 바꾸고 결과를 관찰하자.',variant_files={'main.c':'''#include <stdio.h>\n#define MAX_MACRO(a, b) ((a) > (b) ? (a) : (b))\nint main(void)\n{\n    int count = 4;\n    int result = MAX_MACRO(count, 3);\n    printf("count=%d result=%d\\n", count, result);\n    return 0;\n}\n'''},variant_expected='count=4 result=4\n',question='다중 평가 문제를 괄호를 더 넣어 해결할 수 있는가?',answer='아니다. 괄호는 묶음을 바꿀 뿐 인수가 본문에 등장하는 횟수를 줄이지 못한다.'),
 dict(id='30c',level='부분 완성 예제',title='인수를 한 번만 평가하는 일반 함수',why='매크로 대신 함수를 쓰면 인수 표현식은 함수 호출 전에 한 번 평가되고, 함수 본문은 전달받은 값을 다룬다.',files={'main.c':'''#include <stdio.h>\n\nstatic int larger(int a, int b)\n{\n    return a > b ? a : b;\n}\n\nint main(void)\n{\n    int count = 4;\n    int result = larger(++count, 3);\n    printf("count=%d result=%d\\n", count, result);\n    return 0;\n}\n'''},expected='count=5 result=5\n',predict='이전 매크로 예제와 count가 다른 이유는 무엇인가?',explain=['static int larger(...)는 이 번역 단위 안에서만 연결되는 함수를 정의한다. static은 이 함수의 내부 연결을 지정하며, 매크로와 무관하다.','호출 시 ++count가 한 번 평가되어 값 5가 첫 매개변수 a에 전달된다.','함수 안에서는 a와 b라는 별도의 int 매개변수로 비교한다.','return a > b ? a : b는 큰 값을 반환한다. 함수가 호출자의 count를 추가로 증가시키지 않는다.'],trace='count=4 → 호출 인수 ++count 평가로 count=5 → a=5, b=3 → return 5 → result=5.',mistake='함수 인수가 평가되는 순서는 여러 인수 사이에서 일반적으로 보장되지 않는다. 이 예제에는 부작용 있는 인수가 하나뿐이므로 문제가 없다.',variant='초기 count를 2로 바꾸고 같은 함수를 호출해 보자.',variant_files={'main.c':'''#include <stdio.h>\nstatic int larger(int a, int b)\n{\n    return a > b ? a : b;\n}\nint main(void)\n{\n    int count = 2;\n    int result = larger(++count, 3);\n    printf("count=%d result=%d\\n", count, result);\n    return 0;\n}\n'''},variant_expected='count=3 result=3\n',question='static 함수는 호출할 때마다 지역 변수가 보존된다는 뜻인가?',answer='아니다. 함수 선언의 static은 내부 연결을 뜻한다. 블록 안 지역 객체의 static은 저장 기간을 뜻하므로 위치에 따라 의미가 다르다.'),
 dict(id='30d',level='독립 작성 예제',title='조건을 매크로로 읽되 입력값은 단순하게 유지',why='조건의 의도를 명확히 하고 평가 부작용을 제거하여 안전한 코드를 작성한다.',files={'main.c':'''#include <stdio.h>\n#define IN_RANGE(x, low, high) (((x) >= (low)) && ((x) <= (high)))\n\nint main(void)\n{\n    int score = 78;\n    if (IN_RANGE(score, 0, 100)) {\n        printf("valid\\n");\n    } else {\n        printf("invalid\\n");\n    }\n    return 0;\n}\n'''},expected='valid\n',predict='score=78일 때 &&의 두 조건은 각각 무엇을 평가하는가?',explain=['IN_RANGE는 x가 low 이상이면서 high 이하인지 검사하는 두 비교를 &&로 연결한다.','&&는 왼쪽 조건이 거짓이면 오른쪽 조건을 평가하지 않는 단락 평가를 수행한다.','score는 증가·감소가 없는 단순한 지역 객체이므로 두 번 읽혀도 값이 바뀌지 않는다.','if는 매크로 확장 결과가 참일 때 valid를 출력한다.'],trace='score=78 → 78>=0 참 → 78<=100 참 → 참 && 참 → valid 출력.',mistake='IN_RANGE(++score, 0, 100)처럼 인수에 부작용을 넣으면 의도하지 않은 여러 평가가 일어날 수 있다.',variant='score를 105로 바꾸어 invalid를 출력하도록 독립 작성하자.',variant_files={'main.c':'''#include <stdio.h>\n#define IN_RANGE(x, low, high) (((x) >= (low)) && ((x) <= (high)))\nint main(void)\n{\n    int score = 105;\n    if (IN_RANGE(score, 0, 100)) {\n        printf("valid\\n");\n    } else {\n        printf("invalid\\n");\n    }\n    return 0;\n}\n'''},variant_expected='invalid\n',question='&& 오른쪽 비교가 언제나 실행되는가?',answer='아니다. 왼쪽이 거짓이면 오른쪽은 실행되지 않는다. 이것은 C 실행 시점의 단락 평가 규칙이다.')
 ]),
 dict(num='1.31',title='조건부 컴파일과 헤더 가드: 소스가 컴파일 전에 갈라지는 길',
 intro=[
 '조건부 컴파일은 C 실행문 if와 전혀 다른 시점에 작동한다. #if, #ifdef, #ifndef 등은 전처리 단계에서 어떤 토큰을 컴파일러에 전달할지 선택한다. 실행 파일을 만든 뒤 사용자 입력에 따라 갈라지는 if와 혼동하지 않아야 한다.',
 '전처리 조건식에는 정수 전처리 상수식 등 허용된 형식만 사용할 수 있다. #ifdef NAME은 NAME이라는 매크로가 정의됐는지 묻는다. 값이 0인지와는 다른 질문이다. #if NAME은 NAME이 확장되어 조건식에서 0인지 아닌지를 판단한다.',
 '헤더 가드는 같은 헤더가 여러 경로로 포함되더라도 동일한 선언이 한 번만 들어오도록 하는 관용적인 보호 구조다. #ifndef 고유한_이름 → #define 고유한_이름 → 헤더 내용 → #endif의 형태로 구성한다. 실제로는 헤더에 중복 정의가 들어가지 않도록 설계하고, 구현별 #pragma once는 ISO C11 표준 지시문이 아님을 알아야 한다.',
 '여러 소스 파일을 함께 컴파일할 때 헤더에는 선언을, .c 파일에는 정의를 두는 것이 기본이다. 단, 헤더 가드가 외부 연결 함수의 정의를 여러 번 넣는 실수를 자동으로 고쳐 주지는 않는다. 링크 단계에서 중복 정의가 생길 수 있다.'
 ],terms=[('조건부 컴파일','전처리 단계에서 컴파일에 포함할 소스 토큰을 선택하는 기능이다.'),('defined','전처리 조건식에서 매크로 정의 여부를 검사하는 연산자다.'),('헤더 가드','한 번 포함된 헤더 내용을 동일 번역 단위에 반복 포함하지 않도록 보호하는 전처리 구조다.'),('다중 소스 빌드','여러 .c 파일을 각각 번역하여 링커가 함께 연결하는 과정이다.')],
 examples=[
 dict(id='31a',level='완성 예제',title='#if는 실행 중 if가 아니다',why='기능을 컴파일 시점에 선택할 때 조건부 컴파일을 사용한다. 실행 중 조건과 구별하면 진단 메시지와 결과가 이해된다.',files={'main.c':'''#include <stdio.h>\n#define MODE 1\n\nint main(void)\n{\n#if MODE == 1\n    printf("mode=fast\\n");\n#else\n    printf("mode=normal\\n");\n#endif\n    return 0;\n}\n'''},expected='mode=fast\n',predict='MODE가 1이면 다른 printf 문은 실행될까, 아예 번역 대상에서 빠질까?',explain=['#define MODE 1은 MODE를 1로 확장한다.','#if MODE == 1은 전처리 단계에서 정수 조건을 판정한다.','조건이 참이면 첫 번째 printf가 남고 #else 분기의 printf는 번역 대상에서 제외된다.','#endif는 조건부 전처리 구역을 닫는다.'],trace='전처리: MODE→1, 1==1 참 → 첫 번째 printf만 남음 → 컴파일 및 실행: mode=fast 한 줄 출력.',mistake='#if는 실행 중 사용자 입력에 따라 달라지지 않는다. 사용자 입력 분기는 if문을 사용한다.',variant='MODE를 0으로 바꾸어 normal을 출력하자.',variant_files={'main.c':'''#include <stdio.h>\n#define MODE 0\nint main(void)\n{\n#if MODE == 1\n    printf("mode=fast\\n");\n#else\n    printf("mode=normal\\n");\n#endif\n    return 0;\n}\n'''},variant_expected='mode=normal\n',question='전처리로 제외된 코드도 실행 파일 안에 반드시 남는가?',answer='아니다. 제외된 전처리 분기는 컴파일러에 전달되는 번역 결과에 포함되지 않는다.'),
 dict(id='31b',level='완성 예제',title='#ifdef는 정의 여부를 검사한다',why='매크로 값이 0이어도 정의되어 있으면 #ifdef는 참이라는 사실을 명확히 한다.',files={'main.c':'''#include <stdio.h>\n#define TRACE 0\n\nint main(void)\n{\n#ifdef TRACE\n    printf("trace defined\\n");\n#else\n    printf("trace absent\\n");\n#endif\n    return 0;\n}\n'''},expected='trace defined\n',predict='TRACE의 값이 0인데 왜 defined가 출력되는가?',explain=['#define TRACE 0은 TRACE를 정의하면서 확장값을 0으로 둔다.','#ifdef TRACE는 값의 참·거짓이 아니라 이름이 정의됐는지를 검사한다.','정의되어 있으므로 첫 분기가 전처리 결과에 남는다.','값을 검사하고 싶다면 #if TRACE 같은 형태를 사용해야 한다.'],trace='정의 테이블에 TRACE 존재 → #ifdef 참 → 첫 번째 printf만 컴파일 대상 → trace defined 출력.',mistake='#ifdef TRACE와 #if TRACE는 같은 뜻이 아니다. 전자는 정의 여부, 후자는 확장값의 전처리 조건이다.',variant='TRACE 정의를 제거하고 출력이 어떻게 달라지는지 확인하자.',variant_files={'main.c':'''#include <stdio.h>\nint main(void)\n{\n#ifdef TRACE\n    printf("trace defined\\n");\n#else\n    printf("trace absent\\n");\n#endif\n    return 0;\n}\n'''},variant_expected='trace absent\n',question='TRACE를 0으로 정의하면 #ifdef TRACE가 거짓인가?',answer='아니다. 값이 0이어도 이름이 정의됐으면 참이다.'),
 dict(id='31c',level='부분 완성 예제',title='#ifndef와 #define으로 중복 포함 보호하기',why='한 파일에서 같은 헤더가 반복 포함되어도 중복 선언 또는 정의로 인해 문제가 생기지 않도록 가드를 사용한다.',files={'range.h':'''#ifndef RANGE_H_INCLUDED\n#define RANGE_H_INCLUDED\n\nint clamp_int(int value, int low, int high);\n\n#endif\n''','range.c':'''#include "range.h"\n\nint clamp_int(int value, int low, int high)\n{\n    if (value < low) {\n        return low;\n    }\n    if (value > high) {\n        return high;\n    }\n    return value;\n}\n''','main.c':'''#include <stdio.h>\n#include "range.h"\n#include "range.h"\n\nint main(void)\n{\n    printf("clamped=%d\\n", clamp_int(105, 0, 100));\n    return 0;\n}\n'''},expected='clamped=100\n',predict='range.h를 두 번 포함했는데도 함수가 한 번만 정의되는 이유는?',explain=['range.h의 #ifndef RANGE_H_INCLUDED는 아직 이름이 정의되지 않았을 때만 헤더 본문을 처리한다.','#define RANGE_H_INCLUDED는 이후 같은 번역 단위에서 중복 포함 시 헤더 본문을 건너뛰게 만든다.','헤더에는 clamp_int의 선언만 있으며, 실제 함수 정의는 range.c에 한 번만 있다.','main.c는 선언을 보고 함수를 호출하고, 링크 단계에서 range.c의 정의와 연결된다.'],trace='main.c 전처리: 첫 include에서 선언 포함·가드 정의 → 두 번째 include에서는 본문 제외. range.c는 별도 번역 단위로 함수 정의. 실행: 105>100이므로 100 반환.',mistake='헤더 가드는 여러 번 포함을 막는 것이지 여러 .c 파일에 같은 함수 정의를 넣는 것을 허용하는 장치가 아니다.',variant='clamp_int(-5, 0, 100)을 호출해 아래쪽 경계를 확인하자.',variant_files={'range.h':'''#ifndef RANGE_H_INCLUDED\n#define RANGE_H_INCLUDED\nint clamp_int(int value, int low, int high);\n#endif\n''','range.c':'''#include "range.h"\nint clamp_int(int value, int low, int high)\n{\n    if (value < low) { return low; }\n    if (value > high) { return high; }\n    return value;\n}\n''','main.c':'''#include <stdio.h>\n#include "range.h"\nint main(void)\n{\n    printf("clamped=%d\\n", clamp_int(-5, 0, 100));\n    return 0;\n}\n'''},variant_expected='clamped=0\n',question='range.c가 따로 있는데 main.c만 컴파일하면 반드시 연결되는가?',answer='아니다. range.c도 함께 빌드하여 함수 정의가 링크에 제공되도록 해야 한다.'),
 dict(id='31d',level='독립 작성 예제',title='정의 여부와 수치 조건을 결합해 기능 선택',why='빌드 설정에 따라 진단 출력을 포함할지 결정할 때 조건부 컴파일을 활용한다.',files={'main.c':'''#include <stdio.h>\n#define SHOW_DETAILS 1\n\nint main(void)\n{\n    int sum = 2 + 5;\n#if defined(SHOW_DETAILS) && SHOW_DETAILS == 1\n    printf("details: %d + %d\\n", 2, 5);\n#endif\n    printf("sum=%d\\n", sum);\n    return 0;\n}\n'''},expected='details: 2 + 5\nsum=7\n',predict='SHOW_DETAILS가 0이면 sum 출력은 남는가? details 출력은?',explain=['defined(SHOW_DETAILS)는 이름의 정의 여부를 검사한다.','&&는 전처리 조건식에서도 두 조건을 결합한다. SHOW_DETAILS가 1인 경우에만 details 출력문이 남는다.','sum은 실행 중 지역 객체로 2+5를 계산한다.','#endif 바깥의 printf는 설정과 무관하게 항상 컴파일된다.'],trace='전처리: defined 참, 1==1 참 → details 포함. 실행: sum=7 → details 한 줄 → sum 한 줄.',mistake='디버그 출력이 컴파일에서 제외되어도 sum 계산 자체는 영향을 받지 않는다. 학습용 확인 출력과 필수 계산을 분리한다.',variant='SHOW_DETAILS를 0으로 바꾸고 sum만 출력하도록 작성하자.',variant_files={'main.c':'''#include <stdio.h>\n#define SHOW_DETAILS 0\nint main(void)\n{\n    int sum = 2 + 5;\n#if defined(SHOW_DETAILS) && SHOW_DETAILS == 1\n    printf("details: %d + %d\\n", 2, 5);\n#endif\n    printf("sum=%d\\n", sum);\n    return 0;\n}\n'''},variant_expected='sum=7\n',question='컴파일에서 제외된 진단 출력이 필수 계산 로직이어야 하는가?',answer='아니다. 핵심 계산은 조건부 진단 출력과 독립적으로 유지해야 한다.')
 ])
]

# C source verification: compile/run each primary and adjacent variation using two compilers.
records=[]
for section in sections:
    for ex in section['examples']:
        for kind, files, expected in [('primary',ex['files'],ex['expected']),('variation',ex['variant_files'],ex['variant_expected'])]:
            name=f"{ex['id']}_{kind}"
            folder=SRC/name; folder.mkdir(exist_ok=True)
            for fname,content in files.items():
                (folder/fname).write_text(content,encoding='utf-8')
            for compiler in ('gcc','clang'):
                output=folder/(name+'_'+compiler)
                args=[compiler,*FLAGS,*[str(folder/f) for f in files if f.endswith('.c')],'-o',str(output)]
                proc=subprocess.run(args,capture_output=True,text=True,timeout=20)
                run=None
                if proc.returncode==0:
                    run=subprocess.run([str(output)],capture_output=True,text=True,timeout=20)
                ok=proc.returncode==0 and run is not None and run.returncode==0 and run.stdout==expected and run.stderr==''
                records.append(dict(example=ex['id'],kind=kind,compiler=compiler,passed=ok,compile_returncode=proc.returncode,compile_stderr=proc.stderr,run_returncode=run.returncode if run else None,stdout=run.stdout if run else None,expected=expected))
                if not ok: print('FAIL',name,compiler,proc.stderr,run.stdout if run else '')
assert all(x['passed'] for x in records)

# DOCX styles
D=Document(); s=D.sections[0]; s.top_margin=Cm(2.0); s.bottom_margin=Cm(1.9); s.left_margin=Cm(2.45); s.right_margin=Cm(2.25)
styles=D.styles; normal=styles['Normal']; normal.font.name='Noto Sans CJK KR'; normal.font.size=Pt(9.5); normal.paragraph_format.space_after=Pt(5); normal.paragraph_format.line_spacing=1.2
for key,size,color in [('Title',19,RGBColor(13,91,97)),('Heading 1',15,RGBColor(13,106,107)),('Heading 2',11.5,RGBColor(23,88,96)),('Heading 3',10,RGBColor(42,99,105))]:
    st=styles[key]; st.font.name='Noto Sans CJK KR'; st.font.size=Pt(size); st.font.bold=True; st.font.color.rgb=color
    st.paragraph_format.keep_with_next=True; st.paragraph_format.space_before=Pt(10); st.paragraph_format.space_after=Pt(4)
if 'Code' not in styles:
    cs=styles.add_style('Code',WD_STYLE_TYPE.PARAGRAPH); cs.font.name='DejaVu Sans Mono'; cs.font.size=Pt(8); cs.paragraph_format.space_after=Pt(0); cs.paragraph_format.line_spacing=1.08
if 'Aside' not in styles:
    a=styles.add_style('Aside',WD_STYLE_TYPE.PARAGRAPH); a.font.name='Noto Sans CJK KR'; a.font.size=Pt(9); a.font.color.rgb=RGBColor(63,86,90); a.paragraph_format.space_after=Pt(6)

def add(t,style=None):
    p=D.add_paragraph(t,style); return p

def heading(t,level=2):
    return D.add_heading(t,level)

def code(content):
    for line in content.rstrip('\n').split('\n'):
        p=D.add_paragraph(style='Code'); p.paragraph_format.keep_together=True
        sh=OxmlElement('w:shd'); sh.set(qn('w:fill'),'F2F7F7'); p._p.get_or_add_pPr().append(sh)
        p.add_run(line if line else ' ')

def show_files(files):
    for name,content in files.items():
        if len(files)>1: add('파일: '+name,'Aside')
        code(content)

D.add_heading('C언어 종합교과서 | 이해 중심 정식 개정판',0)
add('제1장 · 새 정식 본문 1.29–1.31절 | 전처리기의 기초, 함수형 매크로, 조건부 컴파일')
add('이 원고는 이전 거부 초안과 독립적으로 집필한 새 본문이다. 모든 실습은 전역 객체 없이 구성했다. C11을 기본으로 설명하며 구현별 명령행 옵션은 ISO C 언어와 구별한다. 원서의 문장과 문제는 복제하지 않는다.','Aside')
heading('읽는 방법',1)
add('① 맥락을 읽는다. ② 실행 전에 출력을 예측한다. ③ 전체 코드를 실행한다. ④ 문장별 해설과 값 추적을 대조한다. ⑤ 바로 이어지는 변형 실습을 풀고 인접 정답으로 확인한다. ⑥ 다음 절에서 이전 개념을 다시 인출한다.')
add('표준과 도구: 본문 C 프로그램은 ISO C11 기준이다. GCC/Clang의 -std=c11 -Wall -Wextra -Wpedantic -Werror -O0은 컴파일러 명령행 옵션이며 C 문법이 아니다. Windows MSVC에서는 옵션이 다르다.')

for section in sections:
    heading(section['num']+'  '+section['title'],1)
    heading('왜 배우는가: 실제 개발에서 생기는 문제',3)
    for x in section['intro']:add(x)
    heading('새 개념의 정확한 뜻',3)
    for term,definition in section['terms']:add('• '+term+' — '+definition)
    heading('하위 목표 분해: 읽고, 확장하고, 실행하기',3)
    add('첫째, #으로 시작하는 지시문이 번역 시점에 무엇을 바꾸는지 표시한다. 둘째, 전처리 이후 컴파일러가 보게 되는 식과 선언을 추정한다. 셋째, 실행 중 변하는 지역 객체만 별도로 추적한다. 넷째, 출력이 필수 로직인지 학습용 확인인지 구분한다.')
    for ex in section['examples']:
        heading(f"{ex['level']} · {ex['id']}  {ex['title']}",2)
        heading('이 문제가 왜 필요한가 · 언제 쓰는가',3);add(ex['why'])
        heading('실행하기 전에 예측하기',3);add(ex['predict'])
        heading('전체 C11 프로그램 · 편집 가능한 코드',3);show_files(ex['files'])
        heading('문장별 해설: 왜 이 문장을 쓰는가',3)
        for idx,line in enumerate(ex['explain'],1):add(f'{idx}. {line}')
        heading('전처리·실행 흐름과 값 추적',3);add(ex['trace'])
        heading('실제 표준 출력',3);code(ex['expected'])
        heading('오개념 진단과 단계별 디버깅',3)
        add('오개념: '+ex['mistake'])
        add('디버깅 순서: (1) #define과 #if의 철자 및 #endif 짝을 먼저 검사한다. (2) 전처리 후 식을 손으로 펼쳐 본다. (3) 컴파일러 경고를 확인한다. (4) 실제 출력과 예측을 비교해 논리 오류를 수정한다.')
        heading('예제를 바꾸며 실력 다지기 — 바로 이 예제의 변형',3);add(ex['variant'])
        heading('바로 이어지는 모범 정답과 근거',3);show_files(ex['variant_files']);add('정답 출력:');code(ex['variant_expected'])
        add('수정 근거: 기존 예제의 학습 목표는 유지하고, 위 변형 조건에 필요한 선언·식·출력만 변경했다. 원래의 전처리/실행 순서를 다시 대조해 확인한다.')
        heading('스스로 확인하기: 답을 보기 전에 말해 보자',3);add(ex['question']);add('정답과 이유: '+ex['answer'])
    heading('이 절의 누적 인출과 적용',2)
    if section['num']=='1.29':
        add('질문 1. 이전 1.17절의 지역 변수와 DAYS 매크로는 무엇이 다른가? 정답: 지역 변수는 실행 중 저장 객체이고 DAYS는 번역 시점에 토큰으로 확장된다.')
        add('질문 2. printf의 확인 출력 없이 합계 계산만 하려면 무엇을 남기는가? 정답: total 초기화, 반복문, total += day를 남긴다. 출력은 검증을 위한 선택적 수단이다.')
    elif section['num']=='1.30':
        add('질문 1. 1.29절의 #define DAYS 4와 #define SQUARE(x) ((x)*(x))의 차이는? 정답: 전자는 인수 없는 객체형, 후자는 인수 있는 함수형 매크로다.')
        add('질문 2. 값에 의한 전달을 배우던 1.19절과 비교하면 왜 함수가 안전한가? 정답: 호출 인수는 매개변수에 전달되기 전에 평가되고, 함수 본문은 매개변수의 값을 사용하므로 매크로 본문의 반복 치환과 다르다.')
    else:
        add('질문 1. 1.29절의 #include와 1.31절의 헤더 가드는 어떻게 연결되는가? 정답: 포함 과정에서 같은 헤더가 중복 처리될 수 있으므로 가드가 같은 번역 단위의 반복 내용을 막는다.')
        add('질문 2. 1.23절의 외부 연결 개념과 다중 파일의 관계는? 정답: 다른 번역 단위에서 선언을 통해 참조한 함수는 링크 단계에서 하나의 정의와 연결된다.')

heading('장기 누적 복습 · 독립 작성 과제',1)
add('과제 A — 이전 절을 보지 않고 #define LIMIT 5를 이용해 1부터 LIMIT까지의 합을 계산하는 C11 프로그램을 작성하라. 완료 뒤 1.29d의 계산 순서를 응용해 스스로 검사하라.')
add('모범 답안 A:')
code('''#include <stdio.h>\n#define LIMIT 5\nint main(void)\n{\n    int sum = 0;\n    for (int i = 1; i <= LIMIT; ++i) {\n        sum += i;\n    }\n    printf("sum=%d\\n", sum);\n    return 0;\n}''')
add('정답: sum=15. 반복 전 sum=0이고 i가 1,2,3,4,5일 때 각각 더해진다.')
add('과제 B — SQUARE_BAD(1+2)의 값과 이유를 쓰고, 안전한 매크로 정의로 바꿔 보자. 정답: 1+2*1+2=5이며, ((x)*(x))처럼 인수와 전체를 괄호로 감싸면 9가 된다. 단, 부작용 인수에는 여전히 위험하다.')
add('과제 C — #define FLAG 0 아래에서 #ifdef FLAG와 #if FLAG의 참·거짓을 예측하라. 정답: #ifdef는 참, #if는 거짓이다. 정의 여부와 값의 차이다.')
heading('오류 유형별 점검',2)
add('문법·전처리 오류: #endif 누락, 함수형 매크로 이름과 ( 사이 공백, #define 본문 뒤 불필요한 세미콜론을 확인한다. 논리 오류: 매크로 확장식의 우선순위, 인수 중복 평가, #ifdef와 #if의 혼동을 점검한다. 링크 오류: 선언만 있고 정의 파일을 빌드에 포함하지 않았는지 검사한다.')
heading('표준 구분과 앞으로 배울 범위',2)
add('이번 절은 ISO C11 전처리의 #include, #define, #if, #ifdef, #ifndef, #else, #endif, defined의 입문 목표를 다뤘다. #undef, #elif, #error, #line, #pragma, 문자열화 #, 토큰 결합 ##, 표준 사전 정의 매크로, C23 #elifdef/#elifndef/#warning/#embed는 이후 심화 절에서 독립적으로 다룬다. #pragma once는 구현 확장으로 분리한다.')
add('원서 대응: K&R 제4장 4.11 및 부록 A 전처리 영역, Express 제16장 전처리·다중 파일의 기초 학습목표를 독창적 예제에 연결했다. 원서 두 PDF의 모든 LAB/Exercise/Programming/Mini Project/Advanced Topic과 페이지별 전수 대조는 아직 완료되지 않았으므로 해당 항목을 완료로 표시하지 않는다.','Aside')

# prevent orphan headings, add page number footer
footer=s.footer.paragraphs[0]; footer.alignment=WD_ALIGN_PARAGRAPH.CENTER
footer.add_run('C언어 종합교과서 · 이해 중심 정식 개정판 · 1.29–1.31')
filename='C언어_종합교과서_이해중심_정식개정판_제1장_1.29-1.31.docx'
docx=ROOT/filename; D.save(docx)

# verify that every program shown in document matches exact source strings
from docx import Document as ReadDocument
text='\n'.join(p.text for p in ReadDocument(docx).paragraphs)
assert all(ex['title'] in text for sec in sections for ex in sec['examples'])
assert all(ex['expected'].strip().split('\n')[0] in text for sec in sections for ex in sec['examples'])

# ZIP source only, no compiled binaries
zip_path=ROOT/'official_ch01_1_29-1_31_sources.zip'
with zipfile.ZipFile(zip_path,'w',zipfile.ZIP_DEFLATED) as z:
    for p in SRC.rglob('*'):
        if p.is_file() and p.suffix in ('.c','.h'):z.write(p,p.relative_to(ROOT))

now=datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).isoformat(timespec='seconds')
verification={'created_at':now,'section_range':['1.29','1.30','1.31'],'flags':' '.join(FLAGS),'primary_examples':12,'variation_examples':12,'compiler_executions':len(records),'passed':all(x['passed'] for x in records),'records':records,'docx_sha256':hashlib.sha256(docx.read_bytes()).hexdigest(),'source_zip_sha256':hashlib.sha256(zip_path.read_bytes()).hexdigest(),'source_match':'DOCX includes all primary and variation source blocks; source files generated from identical in-memory strings; example titles and outputs rechecked','originals_audit':'NOT COMPLETE: PDF chapter/appendix and per-exercise mapping remains unverified'}
(ROOT/'official_ch01_1_29-1_31_c11_verification.json').write_text(json.dumps(verification,ensure_ascii=False,indent=2),encoding='utf-8')
progress={'current_completed_sections':[f'1.{i}' for i in range(1,32)],'last_completed':'1.31','next':'1.32','scope_audit':'pending','overall_percent':None,'github_status':'pending'}
(ROOT/'official_rewrite_progress_1_1-1_31.json').write_text(json.dumps(progress,ensure_ascii=False,indent=2),encoding='utf-8')
objectives={'1.29':{'objectives':['#include 헤더 인터페이스','#define 객체형 매크로','전처리와 실행의 구분','컴파일러 -E의 비표준성'],'source_mapping':'K&R 4.11 / Express 16 chapter-level, page-by-page not verified','status':'new manuscript and C11 tests passed; source exercise audit pending'},'1.30':{'objectives':['함수형 매크로','우선순위 괄호','인수 중복 평가','함수 대안'],'source_mapping':'K&R 4.11 / Express 16 chapter-level, page-by-page not verified','status':'new manuscript and C11 tests passed; source exercise audit pending'},'1.31':{'objectives':['#if/#ifdef/#ifndef/#else/#endif','defined','헤더 가드','다중 소스 컴파일과 연결'],'source_mapping':'K&R 4.5/4.11 / Express 16 chapter-level, page-by-page not verified','status':'new manuscript and C11 tests passed; source exercise audit pending'}}
(ROOT/'official_rewrite_objective_tracking_1_1-1_31.json').write_text(json.dumps({'previous_sections':'1.1–1.28 (prior manifest; not independently re-audited in this run)','new_sections':objectives,'originals_full_audit':False},ensure_ascii=False,indent=2),encoding='utf-8')
manifest={'edition':'이해 중심 정식 개정판','range':'1.29–1.31','generated_at':now,'docx':filename,'sha256':verification['docx_sha256'],'source_zip':zip_path.name,'script':Path(__file__).name,'verification':'official_ch01_1_29-1_31_c11_verification.json','progress':'official_rewrite_progress_1_1-1_31.json','objectives':'official_rewrite_objective_tracking_1_1-1_31.json','originals_audit':'incomplete','github_status':'pending'}
(ROOT/'official_rewrite_manifest_1_29-1_31.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'docx':str(docx),'docx_size':docx.stat().st_size,'programs':24,'compiler_runs':len(records),'all_pass':True,'zip':str(zip_path)},ensure_ascii=False))