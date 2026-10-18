#include <stdio.h>
#include <stddef.h>

typedef struct {
    const char *source;
    size_t next;
    int saved;
    int occupied;
} Reader;

static int getch(Reader *r)
{
    if (r->occupied) {
        r->occupied = 0;
        return r->saved;
    }
    if (r->source[r->next] == '\0') {
        return EOF;
    }
    return (unsigned char)r->source[r->next++];
}

static int ungetch(Reader *r, int ch)
{
    if (r->occupied || ch == EOF) {
        return 0;
    }
    r->saved = ch;
    r->occupied = 1;
    return 1;
}

int main(void)
{
    Reader reader = {"AB", 0, 0, 0};
    int first = getch(&reader);
    printf("first=%c\n", first);
    printf("push=%s\n", ungetch(&reader, first) ? "ok" : "fail");
    printf("push_again=%s\n", ungetch(&reader, 'Q') ? "ok" : "full");
    printf("again=%c\n", getch(&reader));
    printf("next=%c\n", getch(&reader));
    return 0;
}
