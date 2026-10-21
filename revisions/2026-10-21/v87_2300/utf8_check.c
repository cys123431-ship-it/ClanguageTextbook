#include <stdio.h>
#include <string.h>
#include "utf8_core.h"

static void inspect(const char *label, const char *text)
{
    const unsigned char *bytes = (const unsigned char *)text;
    const size_t length = strlen(text);
    size_t offset = 0;
    size_t count = 0;
    unsigned long scalar = 0;

    while (offset < length) {
        const size_t before = offset;
        if (!utf8_next(bytes, length, &offset, &scalar)) {
            printf("%s: invalid at byte %zu\n", label, before);
            return;
        }
        ++count;
    }
    printf("%s: %zu bytes, %zu codepoints\n", label, length, count);
}

int main(void)
{
    inspect("mixed", "A\xEA\xB0\x80!");
    inspect("overlong", "\xE0\x80\x80");
    inspect("truncated", "\xEA\xB0");
    inspect("surrogate", "\xED\xA0\x80");
    inspect("too-high", "\xF4\x90\x80\x80");
    return 0;
}
