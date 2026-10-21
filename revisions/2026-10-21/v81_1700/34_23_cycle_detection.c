#include <stddef.h>
#include <stdio.h>

typedef struct Node {
    int value;
    struct Node *next;
} Node;

static const Node *find_cycle_entry(const Node *head)
{
    const Node *slow = head;
    const Node *fast = head;

    while (fast != NULL && fast->next != NULL) {
        slow = slow->next;
        fast = fast->next->next;
        if (slow == fast) {
            slow = head;
            while (slow != fast) {
                slow = slow->next;
                fast = fast->next;
            }
            return slow;
        }
    }
    return NULL;
}

static size_t cycle_length(const Node *entry)
{
    size_t count = 1;
    const Node *current = entry->next;
    while (current != entry) {
        ++count;
        current = current->next;
    }
    return count;
}

int main(void)
{
    Node nodes[5] = {
        {10, NULL}, {20, NULL}, {30, NULL},
        {40, NULL}, {50, NULL}
    };
    for (size_t i = 0; i + 1 < 5; ++i) {
        nodes[i].next = &nodes[i + 1];
    }
    nodes[4].next = &nodes[2];

    const Node *entry = find_cycle_entry(&nodes[0]);
    if (entry != NULL) {
        printf("cycle entry index=%td, value=%d\n",
               entry - nodes, entry->value);
        printf("cycle length=%zu\n", cycle_length(entry));
    } else {
        puts("no cycle");
    }

    nodes[4].next = NULL;
    entry = find_cycle_entry(&nodes[0]);
    if (entry == NULL) {
        puts("after repair: no cycle");
    } else {
        puts("after repair: cycle remains");
    }
    return 0;
}
