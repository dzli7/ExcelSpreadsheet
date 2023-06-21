from collections import defaultdict, deque


# Class to represent a graph
class Graph:
    def __init__(self, graph, v):
        self.graph = graph  # dictionary containing adjacency List
        self.v = v  # No. of vertices #CHANGE

    # function to add an edge to graph
    def add_edge(self, u, v):
        self.graph[u].append(v)

    def topological_sort(self):
        # Mark all the vertices as not visited
        visited = [False] * self.v
        stack = []
        # Double ended queue for nodes with no incoming edges
        queue = deque()
        # Store number of edges for a node
        num_edges = [0] * self.v
        for i in range(self.v):
            for j in self.graph[i]:
                num_edges[j] += 1
        # Enqueue vertices with no incoming edges
        min_edges = min(num_edges)
        for i in range(self.v):
            if num_edges[i] == min_edges:
                queue.append(i)
        while queue:
            curr = queue.popleft()
            visited[curr] = True
            stack.append(curr)
            # Decrease the edge count of all neighbors
            for neighbor in self.graph[curr]:
                num_edges[neighbor] -= 1
                # Enqueue if a neighbor no incoming edges
                if num_edges[neighbor] == 0:
                    queue.append(neighbor)
        return stack
