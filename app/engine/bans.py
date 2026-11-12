import re
from .pilot150_bans import StructuredBans as AcceptedStructuredBans
from .ban_engine import company_name

EMPLOYER_IDENTITY_REVISION = '2026-10-16-verified-entity-disambiguation-2'
STNG_SOURCES = [
    'https://www.gsprom.ru/company/structure/ao-stroytransneftegaz/?user_lang=ru',
    'https://stng.ru/ru/Pages/o-kompanii',
]
VERIFIED_NONMATCHES = [
    dict(row='USER_TRANSNEFT_GROUP', pattern=r'(?:строй\s*транс\s*нефте\s*газ|stroy\s*trans\s*nefte\s*gaz)(?: stng ru)?',
         reason='Стройтранснефтегаз входит в структуру Газстройпрома; Транснефть указана заказчиком. Совпадение подстроки не доказывает принадлежность к группе.', sources=STNG_SOURCES),
    dict(row=[60,86], pattern=r'санеко',
         reason='АО САНЕКО, ИНН 6315334618, ОГРН 1026300892452 — самостоятельное точное название нефтедобывающей компании Самары, а не опечатка ТАНЕКО.',
         sources=['https://transportrussia.ru/images/vipuski_tr/tr_41-2020.pdf','https://pfo.rosnedra.gov.ru/data/Files/File/6220.pdf']),
    dict(row=277, pattern=r'технопарк импульс',
         reason='ООО Технопарк Импульс, ИНН 5009076042, ОГРН 1105009003890, Домодедово, выпускает навесное оборудование. Отличается от АО Импульс в строке бан-листа.',
         sources=['https://www.impulse.su/contacts/','https://www.impulse.su/vacansies/']),
]

class StructuredBans(AcceptedStructuredBans):
    """Keep the accepted rules, correct a verified company-name collision."""
    def assess(self, card, source_values, text=None):
        result = super().assess(card, source_values, text)
        excluded = []; checks = []
        for finding in result['findings']:
            name = company_name(finding.get('employer', ''))
            # Exact verified entity only. Do not suppress separate employment
            # at Transneft or any other personal/company/industry prohibition.
            match = next((rule for rule in VERIFIED_NONMATCHES if finding['row'] in (rule['row'] if isinstance(rule['row'],list) else [rule['row']]) and re.fullmatch(rule['pattern'],name)),None)
            if match:
                excluded.append(finding)
                checks.append(dict(employer=finding['employer'],rejected_rule=finding['row'],reason=match['reason'],sources=match['sources'],checked_on='2026-10-16',implementation_revision=EMPLOYER_IDENTITY_REVISION))
        if excluded:
            result['findings'] = [x for x in result['findings'] if x not in excluded]
            result['active'] = [x for x in result['active'] if x not in excluded]
            result['ban'] = 'да' if result['active'] else 'нет'
            result['employer_identity_checks'] = checks
        result['employer_identity_revision'] = EMPLOYER_IDENTITY_REVISION
        return result
