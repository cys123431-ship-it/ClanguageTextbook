#include <stdio.h>
#include <stddef.h>

typedef struct {
    const char *source;
    size_t next;
    int pushed[16];
    size_t used;
} Reader;

static int getch(Reader *r)
{
    if (r->used > 0) {
        return r->pushed[--r->used];
    }
    if (r->source[r->next] == '\0') {
        return EOF;
    }
    return (unsigned char)r->source[r->next++];
}

static int ungetch(Reader *r, int ch)
{
    if (ch == EOF || r->used == 16) {
        return 0;
    }
    r->pushed[r->used++] = ch;
    return 1;
}

static int ungets(Reader *r, const char *s)
{
    size_t length = 0;
    while (s[length] != '\0') {
        ++length;
    }
    if (length > 16 - r->used) {
        return 0;
    }
    while (length > 0) {
        --length;
        if (!ungetch(r, (unsigned char)s[length])) {
            return 0;
        }
    }
    return 1;
}

int main(void)
{
    Reader reader = {"tail", 0, {0}, 0};
    if (!ungets(&reader, "Hi!")) {
        fputs("pushback failed\n", stderr);
        return 1;
    }
    int ch;
    while ((ch = getch(&reader)) != EOF) {
        putchar(ch);
    }
    putchar('\n');
    return 0;
}
