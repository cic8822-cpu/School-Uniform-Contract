import type { RouteName } from '../../types'

interface HeaderProps {
  onNavigate: (route: RouteName) => void
}

const NAV_ITEMS: { route: RouteName; label: string }[] = [
  { route: 'contractGuide', label: '계약방법안내' },
  { route: 'procedure', label: '계약절차' },
  { route: 'input', label: '기초자료입력' },
  { route: 'forms', label: '서식함' },
  { route: 'export', label: '내보내기' },
  { route: 'guide', label: '사용안내' },
]

export function Header({ onNavigate }: HeaderProps) {
  return (
    <header>
      <button className="brand" onClick={() => onNavigate('home')}>
        전북 <b>전북특별자치도교육청</b>
      </button>
      <nav aria-label="주 메뉴">
        {NAV_ITEMS.map((item) => (
          <button key={item.route} onClick={() => onNavigate(item.route)}>
            {item.label}
          </button>
        ))}
      </nav>
    </header>
  )
}
