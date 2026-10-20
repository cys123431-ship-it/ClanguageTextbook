#include <limits.h>
#include <locale.h>
#include <stdio.h>
#include <string.h>
#include <wchar.h>

int main(void)
{
    const char input[] = "\xEA\xB0\x80"; /* UTF-8: U+AC00 */
    mbstate_t read_state = {0};
    wchar_t wide = 0;

    if (setlocale(LC_CTYPE, "") == NULL) {
        puts("locale setup failed");
        return 1;
    } else {
        /* The active locale is now selected. */
    }

    size_t used = mbrtowc(&wide, input, strlen(input), &read_state);
    if (used == (size_t)-1) {
        puts("invalid sequence for current locale");
        return 0;
    } else if (used == (size_t)-2) {
        puts("incomplete sequence");
        return 0;
    } else if (used == 0) {
        puts("null character");
        return 0;
    } else {
        /* A complete non-null character was converted. */
    }

    char encoded[MB_LEN_MAX];
    mbstate_t write_state = {0};
    size_t written = wcrtomb(encoded, wide, &write_state);
    if (written == (size_t)-1) {
        puts("encoding failed");
        return 1;
    } else {
        /* The encoded bytes are ready. */
    }

    int same = used == written && memcmp(input, encoded, used) == 0;
    printf("bytes=%zu roundtrip=%s\n", used, same ? "yes" : "no");
    return 0;
}
