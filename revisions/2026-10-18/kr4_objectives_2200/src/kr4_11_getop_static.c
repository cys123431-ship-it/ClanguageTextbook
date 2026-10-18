#include <ctype.h>
#include <stdio.h>

static int next_token(char text[], size_t capacity)
{
    static int pending = EOF;
    int ch = pending;
    size_t length = 0u;
    pending = EOF;
    if (ch == EOF) {
        ch = getchar();
    }
    while (ch != EOF && isspace((unsigned char)ch)) {
        ch = getchar();
    }
    if (ch == EOF) {
        return EOF;
    }
    if (!isdigit((unsigned char)ch)) {
        text[0] = (char)ch;
        text[1] = '\0';
        return (int)'O';
    }
    do {
        if (length + 1u >= capacity) {
            do {
                ch = getchar();
            } while (ch != EOF && isdigit((unsigned char)ch));
            if (ch != EOF) {
                pending = ch;
            }
            return (int)'E';
        }
        text[length++] = (char)ch;
        ch = getchar();
    } while (ch != EOF && isdigit((unsigned char)ch));
    text[length] = '\0';
    if (ch != EOF) {
        pending = ch;
    }
    return (int)'N';
}

int main(void)
{
    char text[16];
    int kind;
    while ((kind = next_token(text, sizeof text)) != EOF) {
        if (kind == 'N') {
            printf("NUM:%s\n", text);
        } else if (kind == 'O') {
            printf("OP:%s\n", text);
        } else {
            puts("ERROR: number too long");
        }
    }
    return 0;
}
