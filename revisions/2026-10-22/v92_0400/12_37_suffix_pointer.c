#include <stdbool.h>
#include <stdio.h>

/* 유효한 널 종료 문자열 두 개를 입력받는 계약이다. */
static bool ends_with(const char *text, const char *suffix)
{
    const char *end = text;
    const char *tail = suffix;

    if (text == NULL || suffix == NULL) {
        return false;
    }
    while (*end != '\0') {
        ++end;
    }
    while (*tail != '\0') {
        ++tail;
    }
    while (tail != suffix) {
        if (end == text) {
            return false;
        }
        --end;
        --tail;
        if (*end != *tail) {
            return false;
        }
    }
    return true;
}

int main(void)
{
    const char *words[] = {"compiler", "compiler", "C", "", "hello"};
    const char *suffixes[] = {"iler", "compiler!", "long", "", ""};
    size_t i;

    for (i = 0; i < sizeof words / sizeof words[0]; ++i) {
        printf("case[%zu]: %s\n", i,
               ends_with(words[i], suffixes[i]) ? "yes" : "no");
    }
    return 0;
}
