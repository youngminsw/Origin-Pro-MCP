"""Quick test: verify server.py can start and list tools."""
import sys
sys.path.insert(0, ".")
from app import mcp
import tools.labtalk
import tools.worksheet
import tools.graph
import tools.style
import tools.fitting
import tools.project

# Try different internal attributes
for attr in ['_tools', '_tool_manager', 'list_tools', '_FastMCP__tools']:
    if hasattr(mcp, attr):
        print(f"Found: mcp.{attr}")
        break
print(f"mcp dir (relevant): {[x for x in dir(mcp) if 'tool' in x.lower()]}")
print("Server loads OK - all tool modules imported without errors")
