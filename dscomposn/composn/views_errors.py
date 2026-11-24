from django.shortcuts import render


def permission_denied(request, exception=None):
    return render(request, 'errors/403.html', {'message': str(exception) if exception else ''}, status=403)
