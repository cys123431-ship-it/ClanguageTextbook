#include <stdio.h>
#include <string.h>
#include "utf8_core.h"

static int prefix_bytes(const unsigned char *bytes, size_t length,
                        size_t wanted, size_t *prefix_length,
                        size_t *total)
{
    size_t offset = 0;
    size_t count = 0;
    size_t cut = 0;
    unsigned long scalar = 0;

    while (offset < length) {
        if (!utf8_next(bytes, length, &offset, &scalar)) {
            return 0;
        }
        ++count;
        if (count == wanted) {
            cut = offset;
        }
    }
    if (wanted > count) {
        cut = length;
    }
    *prefix_length = cut;
    *total = count;
    return 1;
}

int main(void)
{
    const char text[] = "A\xEA\xB0\x80" "B";
    char result[16];
    size_t cut = 0;
    size_t total = 0;
    const size_t wanted = 2;
    const size_t length = strlen(text);

    if (!prefix_bytes((const unsigned char *)text, length, wanted,
                      &cut, &total) || cut >= sizeof result) {
        fputs("invalid UTF-8 or insufficient buffer\n", stderr);
        return 1;
    }
    memcpy(result, text, cut);
    result[cut] = '\0';
    printf("input: %zu bytes, %zu codepoints\n", length, total);
    printf("prefix(2): %s (%zu bytes)\n", result, cut);
    return 0;
}
