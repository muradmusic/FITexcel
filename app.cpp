#ifndef __PROGTEST__
#include <cstdlib>
#include <cstdio>
#include <cstring>
#include <cctype>
#include <climits>
#include <cfloat>
#include <cassert>
#include <cmath>
#include <iostream>
#include <sstream>
#include <fstream>
#include <iomanip>
#include <string>
#include <array>
#include <vector>
#include <list>
#include <set>
#include <map>
#include <stack>
#include <queue>
#include <unordered_set>
#include <unordered_map>
#include <memory>
#include <algorithm>
#include <functional>
#include <iterator>
#include <stdexcept>
#include <variant>
#include <optional>
#include <compare>
#include <charconv>
#include <span>
#include <utility>
#include "expression.h"
using namespace std::literals;
using CValue = std::variant<std::monostate, double, std::string>;

constexpr unsigned SPREADSHEET_CYCLIC_DEPS = 0x01;
// constexpr unsigned                     SPREADSHEET_FUNCTIONS                   = 0x02;
// constexpr unsigned                     SPREADSHEET_FILE_IO                     = 0x04;
// constexpr unsigned                     SPREADSHEET_SPEED                       = 0x08;
// constexpr unsigned                     SPREADSHEET_PARSER                      = 0x10;
#endif /* __PROGTEST__ */

// class CExprBuilder
//{
//   public:
//     virtual void                       opAdd                                   () = 0;
//     virtual void                       opSub                                   () = 0;
//     virtual void                       opMul                                   () = 0;
//     virtual void                       opDiv                                   () = 0;
//     virtual void                       opPow                                   () = 0;
//     virtual void                       opNeg                                   () = 0;
//     virtual void                       opEq                                    () = 0;
//     virtual void                       opNe                                    () = 0;
//     virtual void                       opLt                                    () = 0;
//     virtual void                       opLe                                    () = 0;
//     virtual void                       opGt                                    () = 0;
//     virtual void                       opGe                                    () = 0;
//     virtual void                       valNumber                               ( double                                val ) = 0;
//     virtual void                       valString                               ( std::string                           val ) = 0;
//     virtual void                       valReference                            ( std::string                           val ) = 0;
//     virtual void                       valRange                                ( std::string                           val ) = 0;
//     virtual void                       funcCall                                ( std::string                           fnName,
//                                                                                  int                                   paramCount ) = 0;
// };
//
// void                                   parseExpression                         ( std::string                           expr,
//                                                                                  CExprBuilder                        & builder );

typedef enum : int32_t
{
  C_ADD = 0,
  C_SUB,
  C_MUL,
  C_DIV,
  C_POW,
  C_NEG,
  C_EQ,
  C_NE,
  C_LT,
  C_LE,
  C_GT,
  C_GE
} COp;

class CPos
{
public:
  CPos() = default;
  CPos(std::pair<int, int> pos) : m_pos(pos) {}
  CPos(std::string_view str);
  std::pair<int, int> getPos() const { return m_pos; }

  bool operator<(const CPos &rhs) const { return m_pos < rhs.m_pos; }

public:
  std::pair<int, int> m_pos;
};

struct CRef
{
  CRef() {}
  CRef(CPos pos, std::pair<bool, bool> fix) : m_pos(pos), m_fix(fix) {}
  CRef(std::string &val)
  {
    m_fix.first = (val[0] == '$');
    std::string s = m_fix.first ? val.substr(1, val.size() - 1) : val;
    m_fix.second = false;
    for (size_t i = 0; i < s.size(); ++i)
    {
      if (s[i] == '$')
      {
        m_fix.second = true;
        s.erase(i, 1);
        break;
      }
    }
    m_pos = CPos(s);
  }
  CPos m_pos;
  std::pair<bool, bool> m_fix;
};

bool IsLit(char ch)
{
  return (ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z');
}
int LitIndex(char ch)
{
  return (ch >= 'a' && ch <= 'z') ? (ch - 'a' + 1) : (ch - 'A' + 1);
}

// `str`: A string view containing the cell reference.
CPos::CPos(std::string_view str)
{
  int endLit = 0;
  while (endLit < (int)str.size() && IsLit(str[endLit]))
  {
    endLit++;
  }
  if (endLit == 0 || endLit == (int)str.size())
  {
    throw std::invalid_argument("0 literals or only literals");
  }

  int endNum = endLit;
  while (endNum < (int)str.size() && str[endNum] >= '0' && str[endNum] <= '9')
  {
    endNum++;
  }
  if (endNum != (int)str.size())
  {
    throw std::invalid_argument("some trash in string pos");
  }

  // Convert the column letters to a number. For example, 'A' becomes 1, 'B' becomes 2, ..., 'AA' becomes 27..
  m_pos.first = 0;
  for (int i = 0; i < endLit; ++i)
  {
    m_pos.first *= 26;
    m_pos.first += LitIndex(str[i]);
  }
  m_pos.second = 0;
  for (int i = endLit; i < endNum; ++i)
  {
    m_pos.second *= 10;
    m_pos.second += str[i] - '0';
  }
}

class CEBuilder : public CExprBuilder
{
public:
  using Val = std::variant<std::monostate, double, std::string, CRef, COp>;
  virtual void opAdd() override { m_stack.push_back(COp::C_ADD); }
  virtual void opSub() override { m_stack.push_back(COp::C_SUB); }
  virtual void opMul() override { m_stack.push_back(COp::C_MUL); }
  virtual void opDiv() override { m_stack.push_back(COp::C_DIV); }
  virtual void opPow() override { m_stack.push_back(COp::C_POW); }
  virtual void opNeg() override { m_stack.push_back(COp::C_NEG); }
  virtual void opEq() override { m_stack.push_back(COp::C_EQ); }
  virtual void opNe() override { m_stack.push_back(COp::C_NE); }
  virtual void opLt() override { m_stack.push_back(COp::C_LT); }
  virtual void opLe() override { m_stack.push_back(COp::C_LE); }
  virtual void opGt() override { m_stack.push_back(COp::C_GT); }
  virtual void opGe() override { m_stack.push_back(COp::C_GE); }
  virtual void valNumber(double val) override { m_stack.push_back(val); }
  virtual void valString(std::string val) override { m_stack.push_back(val); }
  virtual void valReference(std::string val) override { m_stack.push_back(CRef(val)); }
  virtual void valRange(std::string val) override {}
  virtual void funcCall(std::string fnName,
                        int paramCount) override {}

  void clear() { m_stack.clear(); }

  bool save(std::ostream &os) const;
  bool load(std::istream &is);

  std::vector<Val> m_stack;
};

class CSpreadsheet
{
public:
  static unsigned capabilities()
  {
    return SPREADSHEET_CYCLIC_DEPS;
    //    return SPREADSHEET_CYCLIC_DEPS | SPREADSHEET_FUNCTIONS | SPREADSHEET_FILE_IO | SPREADSHEET_SPEED | SPREADSHEET_PARSER;
  }
  CSpreadsheet() {}
  bool load(std::istream &is);
  bool save(std::ostream &os) const;
  bool setCell(CPos pos, std::string contents);
  CValue getValue(CPos pos);
  void copyRect(CPos dst, CPos src, int w = 1, int h = 1);

private:
  CValue getValue(CPos pos, std::set<CPos> &deps);
  CValue calc(std::vector<CEBuilder::Val> &stack, std::set<CPos> &deps);

  std::map<std::pair<int, int>, CEBuilder> m_tab;
};

bool CEBuilder::save(std::ostream &os) const
{
  os << m_stack.size() << ' ';
  // std::monostate, double, std::string, CRef, COp
  for (size_t i = 0; i < m_stack.size(); ++i)
  {
    os << m_stack[i].index() << ' ';
    if (m_stack[i].index() == 1)
    {
      os << std::get<double>(m_stack[i]) << ' ';
    }
    else if (m_stack[i].index() == 2)
    {
      auto s = std::get<std::string>(m_stack[i]);

      os << s.size() << ' ' << s << ' ';
    }
    else if (m_stack[i].index() == 3)
    {
      auto ref = std::get<CRef>(m_stack[i]);

      os << ref.m_pos.m_pos.first << ' ' << ref.m_pos.m_pos.second << ' ' << (int)(ref.m_fix.first ? 1 : 0) << ' ' << (int)(ref.m_fix.second ? 1 : 0) << ' ';
    }
    else
    {
      auto op = std::get<COp>(m_stack[i]);
      os << int(op) << ' ';
    }
  }

  return true;
}

bool CEBuilder::load(std::istream &is)
{
  m_stack.clear();

  size_t size;
  if (!(is >> size))
    return false;
  // std::monostate, double, std::string, CRef, COp
  for (size_t i = 0; i < size; ++i)
  {
    size_t idx;
    if (!(is >> idx))
      return false;
    if (idx > 4)
    {
      return false;
    }

    if (idx == 0)
    {
      m_stack.push_back(Val());
    }
    else if (idx == 1)
    {
      double tmp;

      if (!(is >> tmp))
        return false;

      m_stack.push_back(tmp);
    }
    else if (idx == 2)
    {
      size_t sSize;

      if (!(is >> sSize))
        return false;

      if (!is.get())
      { // read ' '
        return false;
      }

      std::string tmp;
      for (size_t i = 0; i < sSize; ++i)
      {
        char ch;
        if (!(ch = is.get()))
        { // read ' '
          return false;
        }
        tmp.push_back(ch);
      }

      m_stack.push_back(tmp);
    }
    else if (idx == 3)
    {
      CRef tmp;

      if (!(is >> tmp.m_pos.m_pos.first >> tmp.m_pos.m_pos.second))
        return false;

      std::pair<int, int> fix;

      if (!(is >> fix.first >> fix.second))
        return false;

      tmp.m_fix.first = fix.first;
      tmp.m_fix.second = fix.second;

      m_stack.push_back(tmp);
    }
    else
    {
      int tmp;
      if (!(is >> tmp))
        return false;

      if (tmp < 0 || tmp > 11)
      {
        return false;
      }

      m_stack.push_back(COp(tmp));
    }
  }

  return true;
}
// Saves the current state of the spreadsheet to an output stream.
bool CSpreadsheet::save(std::ostream &os) const
{
  os << m_tab.size() << ' ';
  for (const auto &p : m_tab)
  {
    os << p.first.first << ' ' << p.first.second << ' ';
    p.second.save(os);
  }

  return true;
}
// Loads the state of the spreadsheet from an input stream,
bool CSpreadsheet::load(std::istream &is)
{
  m_tab.clear();
  size_t size;
  if (!(is >> size))
    return false;

  for (size_t i = 0; i < size; ++i)
  {
    std::pair<int, int> pos;

    if (!(is >> pos.first >> pos.second))
      return false;
    CEBuilder tmp;
    if (!tmp.load(is))
      return false;

    m_tab[pos] = tmp;
  }

  return true;
}
// Copy a rectangular block of cells from a source position to a destination within the spreadsheet.
void CSpreadsheet::copyRect(CPos dst, CPos src, int w, int h)
{
  auto sPos = src.getPos();
  auto dPos = dst.getPos();
  std::vector<std::pair<CPos, CEBuilder>> cpy;
  for (int i = 0; i < w; ++i)
  {
    for (int j = 0; j < h; ++j)
    {
      std::pair<int, int> pp = {sPos.first + i, sPos.second + j};
      if (m_tab.find(pp) != m_tab.end())
      {
        cpy.push_back({CPos(pp), m_tab[pp]});
      }
    }
  }

  std::pair<int, int> diff = {dPos.first - sPos.first, dPos.second - sPos.second};

  for (size_t i = 0; i < cpy.size(); ++i)
  {
    CPos pos = cpy[i].first;
    pos.m_pos.first += diff.first;

    pos.m_pos.second += diff.second;
    CEBuilder tmp = cpy[i].second;

    for (int i = 0; i < (int)tmp.m_stack.size(); ++i)
    {
      if (std::holds_alternative<CRef>(tmp.m_stack[i]))
      {
        CRef ref = std::get<CRef>(tmp.m_stack[i]);
        if (!ref.m_fix.first)
        {
          ref.m_pos.m_pos.first += diff.first;
        }
        if (!ref.m_fix.second)
        {
          ref.m_pos.m_pos.second += diff.second;
        }
        //        e = CEBuilder::Val(ref);
        tmp.m_stack[i] = ref;
      }
    }

    m_tab[pos.m_pos] = tmp;
  }
}

bool CSpreadsheet::setCell(CPos pos, std::string contents)
{
  CEBuilder tmp;
  parseExpression(contents, tmp);
  m_tab[pos.getPos()] = tmp;
  return true;
}

CValue CSpreadsheet::getValue(CPos pos)
{
  std::set<CPos> tmp;
  return getValue(pos, tmp);
}

CValue CSpreadsheet::getValue(CPos pos, std::set<CPos> &deps)
{
  if (deps.find(pos) != deps.end() || m_tab.find(pos.getPos()) == m_tab.end())
  {
    return CValue();
  }
  deps.insert(pos);
  auto stack = m_tab[pos.getPos()].m_stack;

  CValue res = calc(stack, deps);

  deps.erase(pos);

  if (stack.size() != 0)
  {
    return CValue();
  }
  return res;
}

// `stack`: A vector of values representing the postfix expression to evaluate.
// `deps`: A set of positions tracking which cells are currently being evaluated to detect circular dependencies.
CValue CSpreadsheet::calc(std::vector<CEBuilder::Val> &stack, std::set<CPos> &deps)
{
  if (stack.size() == 0)
  {
    return CValue();
  }
  auto top = stack[stack.size() - 1];
  stack.pop_back();

  // Process the element based on its type:
  // std::monostate, double, std::string, CRef, COp
  if (top.index() == 0)
  {
    return CValue();
  }
  else if (top.index() == 1)
  {
    return std::get<double>(top);
  }
  else if (top.index() == 2)
  {
    return std::get<std::string>(top);
  }
  else if (top.index() == 3)
  {
    auto tmp = std::get<CRef>(top);
    return getValue(tmp.m_pos, deps);
  }
  else if (top.index() == 4)
  {
    auto op = std::get<COp>(top);
    if (op == COp::C_NEG)
    {
      auto r = calc(stack, deps);
      if (std::holds_alternative<double>(r))
      {
        return -std::get<double>(r);
      }
      return CValue();
    }
    else
    {
      auto r = calc(stack, deps);
      auto l = calc(stack, deps);

      if (r.index() == 0 || l.index() == 0)
      {
        return CValue();
      }

      switch (op)
      {
      case C_ADD:
      {
        if (std::holds_alternative<double>(r) && std::holds_alternative<double>(l))
        {
          return std::get<double>(r) + std::get<double>(l);
        }
        return ((std::holds_alternative<double>(l) ? std::to_string(std::get<double>(l)) : std::get<std::string>(l)) + (std::holds_alternative<double>(r) ? std::to_string(std::get<double>(r)) : std::get<std::string>(r)));
        break;
      }
      case C_SUB:
      {
        if (std::holds_alternative<double>(r) && std::holds_alternative<double>(l))
        {
          return std::get<double>(l) - std::get<double>(r);
        }
        return CValue();
        break;
      }
      case C_MUL:
      {
        if (std::holds_alternative<double>(r) && std::holds_alternative<double>(l))
        {
          return std::get<double>(r) * std::get<double>(l);
        }
        return CValue();
        break;
      }
      case C_DIV:
      {
        if (std::holds_alternative<double>(r) && std::holds_alternative<double>(l))
        {
          double rr = std::get<double>(r);
          if (rr == 0)
          {
            return CValue();
          }
          return std::get<double>(l) / rr;
        }
        return CValue();
        break;
      }
      case C_POW:
      {
        if (std::holds_alternative<double>(r) && std::holds_alternative<double>(l))
        {
          return std::pow(std::get<double>(l), std::get<double>(r));
        }
        return CValue();
        break;
      }
      case C_EQ:
      {
        if (std::holds_alternative<double>(r) && std::holds_alternative<double>(l))
        {
          return (std::get<double>(l) == std::get<double>(r) ? 1. : 0.);
        }
        if (std::holds_alternative<std::string>(r) && std::holds_alternative<std::string>(l))
        {
          return (std::get<std::string>(l) == std::get<std::string>(r) ? 1. : 0.);
        }
        return CValue();
        break;
      }
      case C_NE:
      {
        if (std::holds_alternative<double>(r) && std::holds_alternative<double>(l))
        {
          return (std::get<double>(l) != std::get<double>(r) ? 1. : 0.);
        }
        if (std::holds_alternative<std::string>(r) && std::holds_alternative<std::string>(l))
        {
          return (std::get<std::string>(l) != std::get<std::string>(r) ? 1. : 0.);
        }
        return CValue();
        break;
      }
      case C_LT:
      {
        if (std::holds_alternative<double>(r) && std::holds_alternative<double>(l))
        {
          return (std::get<double>(l) < std::get<double>(r) ? 1. : 0.);
        }
        if (std::holds_alternative<std::string>(r) && std::holds_alternative<std::string>(l))
        {
          return (std::get<std::string>(l) < std::get<std::string>(r) ? 1. : 0.);
        }
        return CValue();
        break;
      }
      case C_LE:
      {
        if (std::holds_alternative<double>(r) && std::holds_alternative<double>(l))
        {
          return (std::get<double>(l) <= std::get<double>(r) ? 1. : 0.);
        }
        if (std::holds_alternative<std::string>(r) && std::holds_alternative<std::string>(l))
        {
          return (std::get<std::string>(l) <= std::get<std::string>(r) ? 1. : 0.);
        }
        return CValue();
        break;
      }
      case C_GT:
      {
        if (std::holds_alternative<double>(r) && std::holds_alternative<double>(l))
        {
          return (std::get<double>(l) > std::get<double>(r) ? 1. : 0.);
        }
        if (std::holds_alternative<std::string>(r) && std::holds_alternative<std::string>(l))
        {
          return (std::get<std::string>(l) > std::get<std::string>(r) ? 1. : 0.);
        }
        return CValue();
        break;
      }
      case C_GE:
      {
        if (std::holds_alternative<double>(r) && std::holds_alternative<double>(l))
        {
          return (std::get<double>(l) >= std::get<double>(r) ? 1. : 0.);
        }
        if (std::holds_alternative<std::string>(r) && std::holds_alternative<std::string>(l))
        {
          return (std::get<std::string>(l) >= std::get<std::string>(r) ? 1. : 0.);
        }
        return CValue();
        break;
      }

      default:
        return CValue();
        break;
      }
    }
  }

  return CValue();
}

#ifndef __PROGTEST__

bool valueMatch(const CValue &r,
                const CValue &s)

{
  if (r.index() != s.index())
    return false;
  if (r.index() == 0)
    return true;
  if (r.index() == 2)
    return std::get<std::string>(r) == std::get<std::string>(s);
  if (std::isnan(std::get<double>(r)) && std::isnan(std::get<double>(s)))
    return true;
  if (std::isinf(std::get<double>(r)) && std::isinf(std::get<double>(s)))
    return (std::get<double>(r) < 0 && std::get<double>(s) < 0) || (std::get<double>(r) > 0 && std::get<double>(s) > 0);
  return fabs(std::get<double>(r) - std::get<double>(s)) <= 1e8 * DBL_EPSILON * fabs(std::get<double>(r));
}

int main()
{
  CSpreadsheet x0, x1;
  std::ostringstream oss;
  std::istringstream iss;
  std::string data;
  assert(x0.setCell(CPos("A1"), "10"));
  assert(x0.setCell(CPos("A2"), "20.5"));
  assert(x0.setCell(CPos("A3"), "3e1"));
  assert(x0.setCell(CPos("A4"), "=40"));
  assert(x0.setCell(CPos("A5"), "=5e+1"));
  assert(x0.setCell(CPos("A6"), "raw text with any characters, including a quote \" or a newline\n"));
  assert(x0.setCell(CPos("A7"), "=\"quoted string, quotes must be doubled: \"\". Moreover, backslashes are needed for C++.\""));
  assert(valueMatch(x0.getValue(CPos("A1")), CValue(10.0)));
  assert(valueMatch(x0.getValue(CPos("A2")), CValue(20.5)));
  assert(valueMatch(x0.getValue(CPos("A3")), CValue(30.0)));
  assert(valueMatch(x0.getValue(CPos("A4")), CValue(40.0)));
  assert(valueMatch(x0.getValue(CPos("A5")), CValue(50.0)));
  assert(valueMatch(x0.getValue(CPos("A6")), CValue("raw text with any characters, including a quote \" or a newline\n")));
  assert(valueMatch(x0.getValue(CPos("A7")), CValue("quoted string, quotes must be doubled: \". Moreover, backslashes are needed for C++.")));
  assert(valueMatch(x0.getValue(CPos("A8")), CValue()));
  assert(valueMatch(x0.getValue(CPos("AAAA9999")), CValue()));
  assert(x0.setCell(CPos("B1"), "=A1+A2*A3"));
  assert(x0.setCell(CPos("B2"), "= -A1 ^ 2 - A2 / 2   "));
  assert(x0.setCell(CPos("B3"), "= 2 ^ $A$1"));
  assert(x0.setCell(CPos("B4"), "=($A1+A$2)^2"));
  assert(x0.setCell(CPos("B5"), "=B1+B2+B3+B4"));
  assert(x0.setCell(CPos("B6"), "=B1+B2+B3+B4+B5"));
  assert(valueMatch(x0.getValue(CPos("B1")), CValue(625.0)));
  assert(valueMatch(x0.getValue(CPos("B2")), CValue(-110.25)));
  assert(valueMatch(x0.getValue(CPos("B3")), CValue(1024.0)));
  assert(valueMatch(x0.getValue(CPos("B4")), CValue(930.25)));
  assert(valueMatch(x0.getValue(CPos("B5")), CValue(2469.0)));
  assert(valueMatch(x0.getValue(CPos("B6")), CValue(4938.0)));
  assert(x0.setCell(CPos("A1"), "12"));
  assert(valueMatch(x0.getValue(CPos("B1")), CValue(627.0)));
  assert(valueMatch(x0.getValue(CPos("B2")), CValue(-154.25)));
  assert(valueMatch(x0.getValue(CPos("B3")), CValue(4096.0)));
  assert(valueMatch(x0.getValue(CPos("B4")), CValue(1056.25)));
  assert(valueMatch(x0.getValue(CPos("B5")), CValue(5625.0)));
  assert(valueMatch(x0.getValue(CPos("B6")), CValue(11250.0)));
  x1 = x0;
  assert(x0.setCell(CPos("A2"), "100"));
  assert(x1.setCell(CPos("A2"), "=A3+A5+A4"));
  assert(valueMatch(x0.getValue(CPos("B1")), CValue(3012.0)));
  assert(valueMatch(x0.getValue(CPos("B2")), CValue(-194.0)));
  assert(valueMatch(x0.getValue(CPos("B3")), CValue(4096.0)));
  assert(valueMatch(x0.getValue(CPos("B4")), CValue(12544.0)));
  assert(valueMatch(x0.getValue(CPos("B5")), CValue(19458.0)));
  assert(valueMatch(x0.getValue(CPos("B6")), CValue(38916.0)));
  assert(valueMatch(x1.getValue(CPos("B1")), CValue(3612.0)));
  assert(valueMatch(x1.getValue(CPos("B2")), CValue(-204.0)));
  assert(valueMatch(x1.getValue(CPos("B3")), CValue(4096.0)));
  assert(valueMatch(x1.getValue(CPos("B4")), CValue(17424.0)));
  assert(valueMatch(x1.getValue(CPos("B5")), CValue(24928.0)));
  assert(valueMatch(x1.getValue(CPos("B6")), CValue(49856.0)));
  oss.clear();
  oss.str("");
  assert(x0.save(oss));
  data = oss.str();
  iss.clear();
  iss.str(data);
  assert(x1.load(iss));
  assert(valueMatch(x1.getValue(CPos("B1")), CValue(3012.0)));
  assert(valueMatch(x1.getValue(CPos("B2")), CValue(-194.0)));
  assert(valueMatch(x1.getValue(CPos("B3")), CValue(4096.0)));
  assert(valueMatch(x1.getValue(CPos("B4")), CValue(12544.0)));
  assert(valueMatch(x1.getValue(CPos("B5")), CValue(19458.0)));
  assert(valueMatch(x1.getValue(CPos("B6")), CValue(38916.0)));
  assert(x0.setCell(CPos("A3"), "4e1"));
  assert(valueMatch(x1.getValue(CPos("B1")), CValue(3012.0)));
  assert(valueMatch(x1.getValue(CPos("B2")), CValue(-194.0)));
  assert(valueMatch(x1.getValue(CPos("B3")), CValue(4096.0)));
  assert(valueMatch(x1.getValue(CPos("B4")), CValue(12544.0)));
  assert(valueMatch(x1.getValue(CPos("B5")), CValue(19458.0)));
  assert(valueMatch(x1.getValue(CPos("B6")), CValue(38916.0)));
  oss.clear();
  oss.str("");
  assert(x0.save(oss));
  data = oss.str();
  for (size_t i = 0; i < std::min<size_t>(data.length(), 10); i++)
    data[i] ^= 0x5a;
  iss.clear();
  iss.str(data);
  assert(!x1.load(iss));
  assert(x0.setCell(CPos("D0"), "10"));
  assert(x0.setCell(CPos("D1"), "20"));
  assert(x0.setCell(CPos("D2"), "30"));
  assert(x0.setCell(CPos("D3"), "40"));
  assert(x0.setCell(CPos("D4"), "50"));
  assert(x0.setCell(CPos("E0"), "60"));
  assert(x0.setCell(CPos("E1"), "70"));
  assert(x0.setCell(CPos("E2"), "80"));
  assert(x0.setCell(CPos("E3"), "90"));
  assert(x0.setCell(CPos("E4"), "100"));
  assert(x0.setCell(CPos("F10"), "=D0+5"));
  assert(x0.setCell(CPos("F11"), "=$D0+5"));
  assert(x0.setCell(CPos("F12"), "=D$0+5"));
  assert(x0.setCell(CPos("F13"), "=$D$0+5"));

  CPos qqq("ABC123");

  x0.copyRect(CPos("G11"), CPos("F10"), 1, 4);
  assert(valueMatch(x0.getValue(CPos("F10")), CValue(15.0)));
  assert(valueMatch(x0.getValue(CPos("F11")), CValue(15.0)));
  assert(valueMatch(x0.getValue(CPos("F12")), CValue(15.0)));
  assert(valueMatch(x0.getValue(CPos("F13")), CValue(15.0)));
  assert(valueMatch(x0.getValue(CPos("F14")), CValue()));
  assert(valueMatch(x0.getValue(CPos("G10")), CValue()));
  assert(valueMatch(x0.getValue(CPos("G11")), CValue(75.0)));
  assert(valueMatch(x0.getValue(CPos("G12")), CValue(25.0)));
  assert(valueMatch(x0.getValue(CPos("G13")), CValue(65.0)));
  assert(valueMatch(x0.getValue(CPos("G14")), CValue(15.0)));
  x0.copyRect(CPos("G11"), CPos("F10"), 2, 4);
  assert(valueMatch(x0.getValue(CPos("F10")), CValue(15.0)));
  assert(valueMatch(x0.getValue(CPos("F11")), CValue(15.0)));
  assert(valueMatch(x0.getValue(CPos("F12")), CValue(15.0)));
  assert(valueMatch(x0.getValue(CPos("F13")), CValue(15.0)));
  assert(valueMatch(x0.getValue(CPos("F14")), CValue()));
  assert(valueMatch(x0.getValue(CPos("G10")), CValue()));
  assert(valueMatch(x0.getValue(CPos("G11")), CValue(75.0)));
  assert(valueMatch(x0.getValue(CPos("G12")), CValue(25.0)));
  assert(valueMatch(x0.getValue(CPos("G13")), CValue(65.0)));
  assert(valueMatch(x0.getValue(CPos("G14")), CValue(15.0)));
  assert(valueMatch(x0.getValue(CPos("H10")), CValue()));
  assert(valueMatch(x0.getValue(CPos("H11")), CValue()));
  assert(valueMatch(x0.getValue(CPos("H12")), CValue()));
  assert(valueMatch(x0.getValue(CPos("H13")), CValue(35.0)));
  assert(valueMatch(x0.getValue(CPos("H14")), CValue()));
  assert(x0.setCell(CPos("F0"), "-27"));
  assert(valueMatch(x0.getValue(CPos("H14")), CValue(-22.0)));
  x0.copyRect(CPos("H12"), CPos("H13"), 1, 2);
  assert(valueMatch(x0.getValue(CPos("H12")), CValue(25.0)));
  assert(valueMatch(x0.getValue(CPos("H13")), CValue(-22.0)));
  assert(valueMatch(x0.getValue(CPos("H14")), CValue(-22.0)));

  return EXIT_SUCCESS;
}
#endif /* __PROGTEST__ */
